import { readFile } from 'node:fs/promises'

import { ZipReader, ZipWriter, BlobReader, BlobWriter, TextReader, Uint8ArrayReader, TextWriter, ZipReaderStream, ZipWriterStream, Uint8ArrayWriter } from '@zip.js/zip.js';
import { DOMParser, Node, XMLSerializer } from '@xmldom/xmldom';

import {traverse} from './DOMUtils.js'

// fillOdtTemplate, getOdtTemplate, getOdtTextContent

/** @typedef {ArrayBuffer} ODTFile */

/**
 * 
 * @param {string} path 
 * @returns {Promise<ODTFile>}
 */
export async function getOdtTemplate(path) {
    const fileBuffer = await readFile(path)
    return fileBuffer.buffer
}


/**
 * @param {ODTFile} odtFile 
 * @returns {Promise<Document>}
 */
async function getContentDocument(odtFile) {
    const reader = new ZipReader(new Uint8ArrayReader(new Uint8Array(odtFile)));

    const entries = await reader.getEntries();

    const contentEntry = entries.find(entry => entry.filename === 'content.xml');

    if (!contentEntry) {
        throw new Error('No content.xml found in the ODT file');
    }

    // @ts-ignore
    const contentText = await contentEntry.getData(new TextWriter());
    await reader.close();

    const parser = new DOMParser();

    return parser.parseFromString(contentText, 'text/xml');
}

/**
 * 
 * @param {Document} odtDocument 
 * @returns {Element}
 */
function getODTTextElement(odtDocument) {
    return odtDocument.getElementsByTagName('office:body')[0]
        .getElementsByTagName('office:text')[0]
}


/**
 * Extracts plain text content from an ODT file, preserving line breaks
 * @param {ArrayBuffer} odtFile - The ODT file as an ArrayBuffer
 * @returns {Promise<string>} Extracted text content
 */
export async function getOdtTextContent(odtFile) {
    const contentDocument = await getContentDocument(odtFile)
    const odtTextElement = getODTTextElement(contentDocument)

    /**
     * 
     * @param {Element} element 
     * @returns {string}
     */
    function getElementTextContent(element){
        //console.log('tagName', element.tagName)
        if(element.tagName === 'text:h' || element.tagName === 'text:p')
            return element.textContent + '\n'
        else{
            const descendantTexts = Array.from(element.childNodes)
                .filter(n => n.nodeType === Node.ELEMENT_NODE)
                .map(getElementTextContent)

            if(element.tagName === 'text:list-item')
                return `- ${descendantTexts.join('')}`

            return descendantTexts.join('')
        }
    }

    return getElementTextContent(odtTextElement)
}




// For a given string, split it into fixed parts and parts to replace

/**
 * @typedef TextPlaceToFill
 * @property { {expression: string, replacedString:string}[] } expressions
 * @property {(values: any) => void} fill
 */


/**
 * @param {string} str
 * @returns {TextPlaceToFill | undefined}
 */
function findPlacesToFillInString(str) {
    const matches = str.matchAll(/\{([^{]+?)\}/g)

    /** @type {TextPlaceToFill['expressions']} */
    const expressions = []

    /** @type {(string | ((data:any) => void))[]} */
    const parts = []
    let remaining = str;

    for (const match of matches) {
        //console.log('match', match)
        const [matched, group1] = match

        const replacedString = matched
        const expression = group1.trim()
        expressions.push({ expression, replacedString })

        const [fixedPart, newRemaining] = remaining.split(replacedString, 2)

        if (fixedPart.length >= 1)
            parts.push(fixedPart)

        // PPP : for now, expression is expected to be only an object property name
        // in the future, it will certainly be a JavaScript expression
        // securely evaluated within an hardernedJS Compartment https://hardenedjs.org/#compartment
        parts.push(data => data[expression])

        remaining = newRemaining
    }

    if (remaining.length >= 1)
        parts.push(remaining)

    //console.log('parts', parts)


    if (remaining === str) {
        // no match found
        return undefined
    }
    else {
        return {
            expressions,
            fill: (data) => {
                return parts.map(p => {
                    if (typeof p === 'string')
                        return p
                    else
                        return p(data)
                })
                    .join('')
            }
        }
    }


}


/**
 * @param {Node} node
 * @returns {TextPlaceToFill[] | undefined}
 */
function findPlacesToFill(node) {
    /** @type {string} */
    let textCandidate

    switch (node.nodeType) {
        case Node.ATTRIBUTE_NODE:
            // @ts-ignore
            textCandidate = node.value

            if (textCandidate) {
                const placesToFill = findPlacesToFillInString(textCandidate)
                return placesToFill ? [{
                    expressions: placesToFill.expressions,
                    fill: data => {
                        node.value = placesToFill.fill(data)
                    }
                }] : undefined
            }

            break;
        case Node.TEXT_NODE:
            // @ts-ignore
            textCandidate = node.data

            if (textCandidate) {
                const placesToFill = findPlacesToFillInString(textCandidate)
                return placesToFill ? [{
                    expressions: placesToFill.expressions,
                    fill: data => {
                        const newText = placesToFill.fill(data)
                        const newTextNode = node.ownerDocument?.createTextNode(newText)
                        node.parentNode?.replaceChild(newTextNode, node)
                    }
                }] : undefined
            }

            break;

        default:
            if (node.childNodes && node.childNodes.length >= 1) {

                return [...node.childNodes]
                    .map(findPlacesToFill)
                    .filter(x => x !== undefined)
                    .flat()
            }

    }

}

/**
 * @param {Element} element 
 */
function findEachElementPair(element){

    /** @type {Element | undefined} */
    let startElement
    /** @type {Element | undefined} */
    let endElement

    let iterableExpression, itemExpression;

    traverse(element, el => {
        //console.log('traverse', el.nodeType, el.nodeName)
        if(Array.from(el.childNodes).some(n => n.nodeType === Node.ELEMENT_NODE)){
            //console.log('skip')
            // skip
        }
        else{
            const text = el.textContent || ''
            //console.log('text', text)

            const eachStartRegex = /{#each\s+([^}]+?)\s+as\s+([^}]+?)\s*}/g;
            const startMatches = [...text.matchAll(eachStartRegex)];

            if(startMatches && startMatches.length >= 1){
                // only consider first match and ignore others for now
                let [_, _iterableExpression, _itemExpression] = startMatches[0]
                
                iterableExpression = _iterableExpression
                itemExpression = _itemExpression
                startElement = el
            }

            const eachEndRegex = /{\/each}/g
            const endMatches = [...text.matchAll(eachEndRegex)];

            if(endMatches && endMatches.length >= 1){
                endElement = el
            }
        }
    })

    //console.log('startElement', startElement)

    if(startElement && endElement){
        return {
            startElement, 
            iterableExpression, 
            itemExpression, 
            endElement
        }
    }
}



/**
 * 
 * @param {Document} contentDocument 
 * @param {any} data 
 */
function fillEachIfExists(contentDocument, data){
    const eachElementPair = findEachElementPair(contentDocument)

    if(eachElementPair){
        const { iterableExpression, itemExpression, startElement, endElement } = eachElementPair

        // find common ancestor
        let commonAncestor

        let startAncestor = startElement
        let endAncestor = endElement
        
        const startAncestry = new Set([startAncestor])
        const endAncestry = new Set([endAncestor]) 

        while(!startAncestry.has(endAncestor) && !endAncestry.has(startAncestor)){
            if(startAncestor.parentNode){
                startAncestor = startAncestor.parentNode
                startAncestry.add(startAncestor)
            }
            if(endAncestor.parentNode){
                endAncestor = endAncestor.parentNode
                endAncestry.add(endAncestor)
            }
        }

        if(startAncestry.has(endAncestor)){
            commonAncestor = endAncestor
        }
        else{
            commonAncestor = startAncestor
        }

        //console.log('commonAncestor', commonAncestor.tagName)
        //console.log('startAncestry', startAncestry.size, [...startAncestry].indexOf(commonAncestor))
        //console.log('endAncestry', endAncestry.size, [...endAncestry].indexOf(commonAncestor))

        const startAncestryToCommonAncestor = [...startAncestry].slice(0, [...startAncestry].indexOf(commonAncestor))
        const endAncestryToCommonAncestor = [...endAncestry].slice(0, [...endAncestry].indexOf(commonAncestor))

        const startChild = startAncestryToCommonAncestor.at(-1)
        const endChild = endAncestryToCommonAncestor.at(-1)

        //console.log('startChild', startChild.tagName)
        //console.log('endChild', endChild.tagName)

        // Find repeatable pattern and extract it in a documentFragment
        const repeatedFragment = contentDocument.createDocumentFragment()

        /** @type {Element[]} */
        const repeatedPatternArray = []
        let sibling = startChild.nextSibling

        while(sibling !== endChild){
            repeatedPatternArray.push(sibling)
            sibling = sibling.nextSibling;
        }

        //console.log('repeatedPatternArray', repeatedPatternArray.length)

        for(const sibling of repeatedPatternArray){
            sibling.parentNode?.removeChild(sibling)
            repeatedFragment.appendChild(sibling)
        }

        // Find the iterable in the data
        // PPP eventually, evaluate the expression as a JS expression
        const iterable = data[iterableExpression]
        if(!iterable){
            throw new TypeError(`Missing iterable (${iterableExpression})`)
        }
        if(typeof iterable[Symbol.iterator] !== 'function'){
            throw new TypeError(`'${iterableExpression}' is not iterable`)
        }

        // create each loop result
        // using a for-of loop to accept all iterable values
        for(const item of iterable){
            const itemFragment = repeatedFragment.cloneNode(true)
            fillTemplateElementWithData(
                // @ts-ignore
                itemFragment, 
                Object.assign({}, data, {[itemExpression]: item})
            )
            commonAncestor.insertBefore(itemFragment, endChild)
        }

        startChild.parentNode.removeChild(startChild)
        endChild.parentNode.removeChild(endChild)
    }
    else{
        // nothing
    }
}


/**
 * 
 * @param {Node} element 
 * @param {any} data 
 * @returns {void}
 */
function fillTemplateElementWithData(element, data){
    // trouver tous les endroits où il y a des choses à remplir
    const placesToFill = findPlacesToFill(element)

    if (placesToFill) {
        //console.log('placesToFill', placesToFill)

        // remplir tous les endroits à remplir
        for (const placeToFill of placesToFill) {
            placeToFill.fill(data)
        }
    }
}



/**
 * 
 * @param {string} contentXml 
 * @param {any} data 
 * @returns {string}
 */
function fillOdtContent(contentXml, data) {
    const parser = new DOMParser();

    const contentDocument = parser.parseFromString(contentXml, 'text/xml');

    const odtTextElement = getODTTextElement(contentDocument)

    fillEachIfExists(contentDocument, data) 

    fillTemplateElementWithData(odtTextElement, data)

    const serializer = new XMLSerializer()

    return serializer.serializeToString(contentDocument)

}


/**
 * @param {ODTFile} odtTemplate
 * @param {any} data 
 * @returns {Promise<ODTFile>}
 */
async function transformOdt(odtTemplate, data) {

    const reader = new ZipReader(new Uint8ArrayReader(new Uint8Array(odtTemplate)));

    // Lire toutes les entrées du fichier ODT
    const entries = reader.getEntriesGenerator();

    // Créer un ZipWriter pour le nouveau fichier ODT
    const writer = new ZipWriter(new Uint8ArrayWriter());

    // Parcourir chaque entrée du fichier ODT
    for await (const entry of entries) {
        //console.log('entry', entry.filename)

        if (entry.filename === "content.xml") {
            // Si l'entrée est content.xml, nous la modifions avec la fonction updateOdtContent
            const contentXml = await entry.getData(new TextWriter());
            
            const updatedContentXml = fillOdtContent(contentXml, data);

            // Ajouter le content.xml modifié au nouveau zip
            await writer.add(entry.filename, new TextReader(updatedContentXml), {
                lastModDate: entry.lastModDate,
                level: 9
            });
        } else {
            // Pour les autres fichiers, les copier tels quels
            const blobWriter = new BlobWriter();
            await entry.getData(blobWriter);
            const blob = await blobWriter.getData();

            // Ajouter l'entrée non modifiée au nouveau zip
            await writer.add(entry.filename, new BlobReader(blob), {
                level: entry.filename === 'mimetype' ? 0 : 9,
                compressionMethod: 0,
                dataDescriptor: false,
                extendedTimestamp: false,
            });
        }
    }

    // Fermer le reader
    await reader.close();

    // Générer et retourner le nouveau fichier ODT
    return writer.close();
}


/**
 * @template T
 * @param {T} data
 * @param {ODTFile} odtTemplate
 * @returns {Promise<ODTFile>}
 */
export async function fillOdtTemplate(odtTemplate, data) {
    return transformOdt(odtTemplate, data)
}








