//@ts-check

import JSZip from 'jszip'

/** @import {SheetName, SheetRawContent, SheetRowRawContent, SheetCellRawContent} from './types.js' */

/**
 * Crée un fichier .ods à partir d'un Map de feuilles de calcul
 * @param {Map<SheetName, SheetRawContent>} sheetsData 
 * @param {(str: string) => Document} parseXML - Function to parse XML content.
 * @param {(doc: Document) => string} serializeXML - Function to parse XML content.
 * @returns {Promise<Uint8Array>}
 */
export async function _createOdsFile(sheetsData, parseXML, serializeXML) {
    const zip = new JSZip();

    // Ajout des fichiers nécessaires au .ods
    zip.file('mimetype', 'application/vnd.oasis.opendocument.spreadsheet');
    
    const contentXml = createContentXml(sheetsData, parseXML, serializeXML);
    zip.file('content.xml', contentXml);
    
    const files = new Map([
        ['/', 'application/vnd.oasis.opendocument.spreadsheet'],
        ['/content.xml', 'text/xml']
    ]);

    const manifestXml = createManifestXml(parseXML, serializeXML, files);
    zip.file('META-INF/manifest.xml', manifestXml);

    // Génération du fichier .ods
    const odsContent = await zip.generateAsync({ type: 'uint8array' });
    return odsContent;
}

/**
 * Crée le contenu XML pour le fichier content.xml
 * @param {Map<SheetName, SheetRawContent>} sheetsData 
 * @param {(str: string) => Document} parseXML - Function to parse XML content.
 * @param {(doc: Document) => string} serializeXML - Function to parse XML content.
 * @returns {string}
 */
function createContentXml(sheetsData, parseXML, serializeXML) {
    const doc = parseXML(
        '<?xml version="1.0" encoding="UTF-8"?><office:document-content></office:document-content>'
    );
    
    const root = doc.documentElement;
    setNamespaces(root);
    
    const body = doc.createElement('office:body');
    const spreadsheet = doc.createElement('office:spreadsheet');
    
    for (const [sheetName, sheetContent] of sheetsData) {
        const table = createTableElement(doc, sheetName, sheetContent);
        spreadsheet.appendChild(table);
    }
    
    body.appendChild(spreadsheet);
    root.appendChild(body);
    
    return serializeXML(doc);
}

/**
 * Crée l'élément XML pour une feuille de calcul
 * @param {Document} doc 
 * @param {string} sheetName 
 * @param {SheetRawContent} sheetContent 
 * @returns {Element}
 */
function createTableElement(doc, sheetName, sheetContent) {
    const table = doc.createElement('table:table');
    table.setAttribute('table:name', sheetName);
    
    sheetContent.forEach((rowContent, rowIndex) => {
        const row = createRowElement(doc, rowContent, rowIndex);
        table.appendChild(row);
    });
    
    return table;
}

/**
 * Crée l'élément XML pour une ligne de la feuille de calcul
 * @param {Document} doc 
 * @param {SheetRowRawContent} rowContent 
 * @param {number} rowIndex 
 * @returns {Element}
 */
function createRowElement(doc, rowContent, rowIndex) {
    const row = doc.createElement('table:table-row');
    
    rowContent.forEach((cell, columnIndex) => {
        const cellElement = createCellElement(doc, cell);
        row.appendChild(cellElement);
    });
    
    return row;
}


/**
 * Convertit les types Excel et ODS en type ODS standard
 * @param {SheetCellRawContent['type']} type 
 * @returns {string}
 */
function mapCellType(type) {
    // Mapping des types Excel vers ODS
    const typeMap = {
        // Types ODS standards
        'float': 'float',
        'percentage': 'percentage',
        'currency': 'currency',
        'date': 'date',
        'time': 'time',
        'boolean': 'boolean',
        'string': 'string',
    };
    
    return typeMap[type] || 'string';
}

/**
 * Crée l'élément XML pour une cellule
 * @param {Document} doc
 * @param {SheetCellRawContent} cell
 * @returns {Element}
 */
function createCellElement(doc, cell) {
    const cellElement = doc.createElement('table:table-cell');
    
    if (cell !== null && cell !== undefined) {
        const { value, type } = cell;
        const cellType = mapCellType(type);
        
        // Définition du type de base
        cellElement.setAttribute('office:value-type', cellType);
        
        // Gestion de la valeur
        if (value !== null && value !== undefined) {
            // Ajout des attributs selon le type converti
            switch (cellType) {
                case 'float':
                case 'percentage':
                    if (!Number.isNaN(Number(value))) {
                        cellElement.setAttribute('office:value', String(value));
                    }
                    break;
                case 'boolean':
                    const boolValue = value.toLowerCase() === 'true' || value === '1';
                    cellElement.setAttribute('office:boolean-value', String(boolValue));
                    break;
                case 'date':
                    const date = new Date(value);
                    if (!Number.isNaN(date.getTime())) {
                        cellElement.setAttribute('office:date-value', 
                            date.toISOString().split('T')[0]);
                    }
                    break;
                case 'time':
                    const time = new Date(`1970-01-01T${value}`);
                    if (!Number.isNaN(time.getTime())) {
                        cellElement.setAttribute('office:time-value', 
                            time.toISOString().split('T')[1].split('.')[0]);
                    }
                    break;
            }
            
            // Ajout du contenu texte
            const textElement = doc.createElement('text:p');
            textElement.textContent = String(value);
            cellElement.appendChild(textElement);
        }
    }
    
    return cellElement;
}


/**
 * Ajoute les espaces de noms nécessaires à l'élément racine
 * @param {Element} root 
 */
function setNamespaces(root) {
    const namespaces = {
        'office': 'urn:oasis:names:tc:opendocument:xmlns:office:1.0',
        'style': 'urn:oasis:names:tc:opendocument:xmlns:style:1.0',
        'text': 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
        'table': 'urn:oasis:names:tc:opendocument:xmlns:table:1.0',
        'draw': 'urn:oasis:names:tc:opendocument:xmlns:drawing:1.0',
        'fo': 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0',
        'xlink': 'http://www.w3.org/1999/xlink',
        'dc': 'http://purl.org/dc/elements/1.1/',
        'meta': 'urn:oasis:names:tc:opendocument:xmlns:meta:1.0',
        'number': 'urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0',
        'presentation': 'urn:oasis:names:tc:opendocument:xmlns:presentation:1.0',
        'svg': 'urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0',
        'chart': 'urn:oasis:names:tc:opendocument:xmlns:chart:1.0',
        'dr3d': 'urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0',
        'math': 'http://www.w3.org/1998/Math/MathML',
        'form': 'urn:oasis:names:tc:opendocument:xmlns:form:1.0',
        'script': 'urn:oasis:names:tc:opendocument:xmlns:script:1.0',
        'ooo': 'http://openoffice.org/2004/office',
        'ooow': 'http://openoffice.org/2004/writer',
        'oooc': 'http://openoffice.org/2004/calc',
        'dom': 'http://www.w3.org/2001/xml-events',
        'xforms': 'http://www.w3.org/2002/xforms',
        'xsd': 'http://www.w3.org/2001/XMLSchema',
        'xsi': 'http://www.w3.org/2001/XMLSchema-instance',
        'rpt': 'http://openoffice.org/2005/report',
        'of': 'urn:oasis:names:tc:opendocument:xmlns:of:1.2',
        'xhtml': 'http://www.w3.org/1999/xhtml',
        'grddl': 'http://www.w3.org/2003/g/data-view#',
        'tableooo': 'http://openoffice.org/2009/table',
        'drawooo': 'http://openoffice.org/2010/draw',
        'calcext': 'urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0',
        'loext': 'urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0',
        'field': 'urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0',
        'formx': 'urn:openoffice:names:experimental:ooxml-odf-interop:xmlns:form:1.0',
        'css3t': 'http://www.w3.org/TR/css3-text/'
    };

    for (const [prefix, uri] of Object.entries(namespaces)) {
        root.setAttribute(`xmlns:${prefix}`, uri);
    }

    root.setAttribute('office:version', '1.2');
}

/**
 * Crée le contenu XML pour le fichier manifest.xml
 * @param {(str: string) => Document} parseXML - Function to parse XML content.
 * @param {(doc: Document) => string} serializeXML - Function to parse XML content.
 * @param {Map<string, string>} files - Map des fichiers avec leur chemin et type MIME
 * @returns {string}
 */
function createManifestXml(parseXML, serializeXML, files) {
    const doc = parseXML(
        '<?xml version="1.0" encoding="UTF-8"?><manifest:manifest></manifest:manifest>'
    );
    const root = doc.documentElement;
    root.setAttribute('manifest:xmlns:manifest', 'urn:oasis:names:tc:opendocument:xmlns:manifest:1.0');

    for(const [path, mediaType] of files){
        const entry = doc.createElement('manifest:file-entry');
        entry.setAttribute('manifest:full-path', path);
        entry.setAttribute('manifest:media-type', mediaType);
        root.appendChild(entry);
    }

    return serializeXML(doc);
}