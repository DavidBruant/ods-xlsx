import { readFile } from 'node:fs/promises'

import { ZipReader, ZipWriter, BlobReader, BlobWriter, TextReader, Uint8ArrayReader, TextWriter, ZipReaderStream, ZipWriterStream } from '@zip.js/zip.js';
import { DOMParser, Node, XMLSerializer } from '@xmldom/xmldom';


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

    const extractedTexts = Array.from(odtTextElement.childNodes)
        .filter(el => {
            if (el.nodeType !== Node.ELEMENT_NODE)
                return false
            else
                // @ts-ignore
                return el.tagName === 'text:h' || el.tagName === 'text:p'
        })
        .map(el => el.textContent)

    // Join paragraphs with newlines to preserve structure
    return extractedTexts.join('\n');
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
 * 
 * @param {string} contentXml 
 * @param {any} data 
 * @returns {string}
 */
function fillOdtContent(contentXml, data) {
    const parser = new DOMParser();

    const contentDocument = parser.parseFromString(contentXml, 'text/xml');

    const odtTextElement = getODTTextElement(contentDocument)

    // trouver tous les endroits où il y a des choses à remplir
    const placesToFill = findPlacesToFill(odtTextElement)

    if (placesToFill) {
        //console.log('placesToFill', placesToFill)

        // remplir tous les endroits à remplir
        for (const placeToFill of placesToFill) {
            placeToFill.fill(data)
        }
    }

    const serializer = new XMLSerializer()

    return serializer.serializeToString(contentDocument)

}

/**
 * 
 * @param {ReadableStream} readableStream 
 * @returns {Promise<ArrayBuffer>}
 */
async function readableStreamToArrayBuffer(readableStream) {
    const reader = readableStream.getReader();
    const chunks = [];

    try {
        while (true) {
            const { done, value } = await reader.read();
            if (done) break;
            chunks.push(value);
        }
    } finally {
        reader.releaseLock();
    }

    // Calculer la taille totale nécessaire
    const totalLength = chunks.reduce((acc, chunk) => acc + chunk.length, 0);

    // Créer un nouveau ArrayBuffer avec la taille totale
    const result = new Uint8Array(totalLength);

    // Copier tous les chunks dans le nouveau ArrayBuffer
    let position = 0;
    for (const chunk of chunks) {
        result.set(chunk, position);
        position += chunk.length;
    }

    return result.buffer;
}


/**
 * @param {ODTFile} odtTemplate
 * @param {any} data 
 * @returns {Promise<ODTFile>}
 */
async function transformOdt(odtTemplate, data) {

    const inputStream = new ReadableStream({
        start(controller) {
            controller.enqueue(new Uint8Array(odtTemplate));
            controller.close();
        }
    });

    // Créer un TransformStream qui modifie content.xml
    const transformStream = new TransformStream({
        async transform(entry, controller) {
            console.log('entry', entry.filename)

            if (entry.filename === "content.xml") {

                // Récupérer le contenu XML
                const contentXmlArrayBuffer = await readableStreamToArrayBuffer(entry.readable);

                const encoder = new TextDecoder("utf-8");

                const contentXml = encoder.decode(contentXmlArrayBuffer)
                // Appliquer les modifications
                const updatedContentXml = fillOdtContent(contentXml, data);

                // Modifier l'entry pour qu'elle utilise le nouveau contenu
                entry.readable = new ReadableStream({
                    start(controller) {
                        const enc = new TextEncoder();
                        controller.enqueue(enc.encode(updatedContentXml));
                        controller.close();
                    }
                });
            }

            // Envoyer l'entry (originale ou modifiée) au controller
            controller.enqueue(entry);
        }
    });

    // new Uint8ArrayReader(new Uint8Array(odtFile))
    const outputZipWriter = new ZipWriterStream()


    await inputStream
        .pipeThrough(new ZipReaderStream())
        .pipeThrough(transformStream)
        .pipeTo(outputZipWriter.writable('blabla.zip'))


    // Récupérer le résultat sous forme de Blob
    return outputZipWriter.close()
}


/**
 * @template T
 * @param {T} data
 * @param {ODTFile} odtTemplate
 * @returns {Promise<ODTFile>}
 */
export async function fillOdtTemplate(odtTemplate, data) {
    const res = await transformOdt(odtTemplate, data)

    console.log('res', res)

    return res;
}








