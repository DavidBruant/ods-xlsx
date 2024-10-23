//@ts-check

import JSZip from 'jszip'

/** @import {SheetName, SheetRawContent, SheetRowRawContent, SheetCellRawContent} from './types.js' */

const officeVersion = '1.2'

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
    
    const stylesXml = createStylesXml(undefined, parseXML, serializeXML);
    zip.file('styles.xml', stylesXml);
    
    const metaXml = createMetaXml(undefined, parseXML, serializeXML);
    zip.file('meta.xml', metaXml);
    
    const settingsXml = createSettingsXml(undefined, parseXML, serializeXML);
    zip.file('settings.xml', settingsXml);
    
    const files = new Map([
        ['/', 'application/vnd.oasis.opendocument.spreadsheet'],
        ['/content.xml', 'text/xml'],
        ['/styles.xml', 'text/xml'],
        ['/meta.xml', 'text/xml'],
        ['/settings.xml', 'text/xml']
    ]);

    const manifestXml = createManifestXml(parseXML, serializeXML, files);
    zip.file('META-INF/manifest.xml', manifestXml, {createFolders: false});

    // Génération du fichier .ods
    const odsContent = await zip.generateAsync({ 
        type: 'uint8array'
    });
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
        `<?xml version="1.0" encoding="utf-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0" xmlns:math="http://www.w3.org/1998/Math/MathML" xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0" xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0" xmlns:ooo="http://openoffice.org/2004/office" xmlns:ooow="http://openoffice.org/2004/writer" xmlns:oooc="http://openoffice.org/2004/calc" xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2" xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0" xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0" xmlns:dom="http://www.w3.org/2001/xml-events" xmlns:xforms="http://www.w3.org/2002/xforms" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" office:version="1.2"><office:font-face-decls><style:font-face style:name="Arial" svg:font-family="Arial" /></office:font-face-decls><office:automatic-styles><style:style style:name="co1" style:family="table-column"><style:table-column-properties fo:break-before="auto" style:column-width="0.8958333134651184in" /></style:style><style:style style:name="co2" style:family="table-column"><style:table-column-properties fo:break-before="auto" style:column-width="0.8958333134651184in" /></style:style><style:style style:name="ro1" style:family="table-row"><style:table-row-properties style:row-height="0.17777777777777778in" fo:break-before="auto" style:use-optimal-row-height="true" /></style:style><style:style style:name="ta1" style:family="table" style:master-page-name="PageStyle_5f_La feuille"><style:table-properties table:display="true" style:writing-mode="lr-tb" /></style:style><style:style style:name="ta2" style:family="table" style:master-page-name="PageStyle_5f_L'autre feuille"><style:table-properties table:display="true" style:writing-mode="lr-tb" /></style:style><style:style style:name="ce1" style:family="table-cell" style:parent-style-name="Default"><style:text-properties  style:font-name="Arial" style:font-name-asian="Arial" fo:font-weight="normal" style:font-weight-asian="normal" style:font-weight-complex="normal" fo:font-style="normal" style:font-style-asian="normal" style:font-style-complex="normal" style:text-line-through-style="none" fo:font-size="10pt" style:font-size-asian="10pt" style:font-size-complex="10pt" fo:color="#000000" style:text-underline-style="none" /></style:style><style:style style:name="ce2" style:family="table-cell" style:parent-style-name="Default"><style:text-properties  style:font-name="Arial" style:font-name-asian="Arial" fo:font-weight="normal" style:font-weight-asian="normal" style:font-weight-complex="normal" fo:font-style="normal" style:font-style-asian="normal" style:font-style-complex="normal" style:text-line-through-style="none" fo:font-size="10pt" style:font-size-asian="10pt" style:font-size-complex="10pt" fo:color="#000000" style:text-underline-style="none" /></style:style><style:style style:name="ce3" style:family="table-cell" style:parent-style-name="Default"><style:text-properties  style:font-name="Arial" style:font-name-asian="Arial" fo:font-weight="normal" style:font-weight-asian="normal" style:font-weight-complex="normal" fo:font-style="normal" style:font-style-asian="normal" style:font-style-complex="normal" style:text-line-through-style="none" fo:font-size="10pt" style:font-size-asian="10pt" style:font-size-complex="10pt" fo:color="#000000" style:text-underline-style="none" /></style:style><style:style style:name="ce4" style:family="table-cell" style:parent-style-name="Default"><style:text-properties  style:font-name="Arial" style:font-name-asian="Arial" fo:font-weight="normal" style:font-weight-asian="normal" style:font-weight-complex="normal" fo:font-style="normal" style:font-style-asian="normal" style:font-style-complex="normal" style:text-line-through-style="none" fo:font-size="10pt" style:font-size-asian="10pt" style:font-size-complex="10pt" fo:color="#000000" style:text-underline-style="none" /></style:style><style:style style:name="ce5" style:family="table-cell" style:parent-style-name="Default" style:data-style-name="N-1"><style:text-properties  style:font-name="Arial" style:font-name-asian="Arial" fo:font-weight="normal" style:font-weight-asian="normal" style:font-weight-complex="normal" fo:font-style="normal" style:font-style-asian="normal" style:font-style-complex="normal" style:text-line-through-style="none" fo:font-size="10pt" style:font-size-asian="10pt" style:font-size-complex="10pt" fo:color="#000000" style:text-underline-style="none" /></style:style><style:style style:name="ce6" style:family="table-cell" style:parent-style-name="Default" style:data-style-name="N-1"><style:text-properties  style:font-name="Arial" style:font-name-asian="Arial" fo:font-weight="normal" style:font-weight-asian="normal" style:font-weight-complex="normal" fo:font-style="normal" style:font-style-asian="normal" style:font-style-complex="normal" style:text-line-through-style="none" fo:font-size="10pt" style:font-size-asian="10pt" style:font-size-complex="10pt" fo:color="#000000" style:text-underline-style="none" /></style:style><style:style style:name="ce7" style:family="table-cell" style:parent-style-name="Default" style:data-style-name="N-1"><style:text-properties  style:font-name="Arial" style:font-name-asian="Arial" fo:font-weight="normal" style:font-weight-asian="normal" style:font-weight-complex="normal" fo:font-style="normal" style:font-style-asian="normal" style:font-style-complex="normal" style:text-line-through-style="none" fo:font-size="10pt" style:font-size-asian="10pt" style:font-size-complex="10pt" fo:color="#000000" style:text-underline-style="none" /></style:style><style:style style:name="ce8" style:family="table-cell" style:parent-style-name="Default" style:data-style-name="N-1"><style:text-properties  style:font-name="Arial" style:font-name-asian="Arial" fo:font-weight="normal" style:font-weight-asian="normal" style:font-weight-complex="normal" fo:font-style="normal" style:font-style-asian="normal" style:font-style-complex="normal" style:text-line-through-style="none" fo:font-size="10pt" style:font-size-asian="10pt" style:font-size-complex="10pt" fo:color="#000000" style:text-underline-style="none" /></style:style><style:style style:name="ce9" style:family="table-cell" style:parent-style-name="Default" style:data-style-name="N-1"><style:text-properties  style:font-name="Arial" style:font-name-asian="Arial" fo:font-weight="normal" style:font-weight-asian="normal" style:font-weight-complex="normal" fo:font-style="normal" style:font-style-asian="normal" style:font-style-complex="normal" style:text-line-through-style="none" fo:font-size="10pt" style:font-size-asian="10pt" style:font-size-complex="10pt" fo:color="#000000" style:text-underline-style="none" /></style:style><style:style style:name="ce10" style:family="table-cell" style:parent-style-name="Default" style:data-style-name="N-1"><style:text-properties  style:font-name="Arial" style:font-name-asian="Arial" fo:font-weight="normal" style:font-weight-asian="normal" style:font-weight-complex="normal" fo:font-style="normal" style:font-style-asian="normal" style:font-style-complex="normal" style:text-line-through-style="none" fo:font-size="10pt" style:font-size-asian="10pt" style:font-size-complex="10pt" fo:color="#000000" style:text-underline-style="none" /></style:style><style:style style:name="ce11" style:family="table-cell" style:parent-style-name="Default" style:data-style-name="N-1"><style:text-properties  style:font-name="Arial" style:font-name-asian="Arial" fo:font-weight="normal" style:font-weight-asian="normal" style:font-weight-complex="normal" fo:font-style="normal" style:font-style-asian="normal" style:font-style-complex="normal" style:text-line-through-style="none" fo:font-size="10pt" style:font-size-asian="10pt" style:font-size-complex="10pt" fo:color="#000000" style:text-underline-style="none" /></style:style><style:style style:name="ce12" style:family="table-cell" style:parent-style-name="Default" style:data-style-name="N-1"><style:text-properties  style:font-name="Arial" style:font-name-asian="Arial" fo:font-weight="normal" style:font-weight-asian="normal" style:font-weight-complex="normal" fo:font-style="normal" style:font-style-asian="normal" style:font-style-complex="normal" style:text-line-through-style="none" fo:font-size="10pt" style:font-size-asian="10pt" style:font-size-complex="10pt" fo:color="#000000" style:text-underline-style="none" /></style:style><style:style style:name="ce13" style:family="table-cell" style:parent-style-name="Default" style:data-style-name="N-1"><style:text-properties  style:font-name="Arial" style:font-name-asian="Arial" fo:font-weight="normal" style:font-weight-asian="normal" style:font-weight-complex="normal" fo:font-style="normal" style:font-style-asian="normal" style:font-style-complex="normal" style:text-line-through-style="none" fo:font-size="10pt" style:font-size-asian="10pt" style:font-size-complex="10pt" fo:color="#000000" style:text-underline-style="none" /></style:style><style:style style:name="ce14" style:family="table-cell" style:parent-style-name="Default" style:data-style-name="N-1"><style:text-properties  style:font-name="Arial" style:font-name-asian="Arial" fo:font-weight="normal" style:font-weight-asian="normal" style:font-weight-complex="normal" fo:font-style="normal" style:font-style-asian="normal" style:font-style-complex="normal" style:text-line-through-style="none" fo:font-size="10pt" style:font-size-asian="10pt" style:font-size-complex="10pt" fo:color="#000000" style:text-underline-style="none" /></style:style><style:style style:name="ce15" style:family="table-cell" style:parent-style-name="Default" style:data-style-name="N-1" /><style:style style:name="T0" style:family="text"><style:text-properties  fo:font-family="Arial" style:font-family-asian="Arial" fo:font-weight="normal" style:font-weight-asian="normal" style:font-weight-complex="normal" fo:font-style="normal" style:font-style-asian="normal" style:font-style-complex="normal" style:text-line-through-style="none" fo:font-size="10pt" style:font-size-asian="10pt" style:font-size-complex="10pt" fo:color="#000000" style:text-underline-style="none" /></style:style><style:style style:name="T1" style:family="text"><style:text-properties  fo:font-family="Arial" style:font-family-asian="Arial" fo:font-weight="normal" style:font-weight-asian="normal" style:font-weight-complex="normal" fo:font-style="normal" style:font-style-asian="normal" style:font-style-complex="normal" style:text-line-through-style="none" fo:font-size="10pt" style:font-size-asian="10pt" style:font-size-complex="10pt" fo:color="#000000" style:text-underline-style="none" /></style:style><style:style style:name="T2" style:family="text"><style:text-properties  fo:font-family="Arial" style:font-family-asian="Arial" fo:font-weight="normal" style:font-weight-asian="normal" style:font-weight-complex="normal" fo:font-style="normal" style:font-style-asian="normal" style:font-style-complex="normal" style:text-line-through-style="none" fo:font-size="10pt" style:font-size-asian="10pt" style:font-size-complex="10pt" fo:color="#000000" style:text-underline-style="none" /></style:style><style:style style:name="T3" style:family="text"><style:text-properties  fo:font-family="Arial" style:font-family-asian="Arial" fo:font-weight="normal" style:font-weight-asian="normal" style:font-weight-complex="normal" fo:font-style="normal" style:font-style-asian="normal" style:font-style-complex="normal" style:text-line-through-style="none" fo:font-size="10pt" style:font-size-asian="10pt" style:font-size-complex="10pt" fo:color="#000000" style:text-underline-style="none" /></style:style><style:style style:name="T4" style:family="text"><style:text-properties  fo:font-family="Arial" style:font-family-asian="Arial" fo:font-weight="normal" style:font-weight-asian="normal" style:font-weight-complex="normal" fo:font-style="normal" style:font-style-asian="normal" style:font-style-complex="normal" style:text-line-through-style="none" fo:font-size="10pt" style:font-size-asian="10pt" style:font-size-complex="10pt" fo:color="#000000" style:text-underline-style="none" /></style:style></office:automatic-styles><office:body><office:spreadsheet><table:calculation-settings table:automatic-find-labels="false" table:use-regular-expressions="false" table:use-wildcards="true" /><table:table table:name="La feuille" table:style-name="ta1" table:print="false"><office:forms form:automatic-focus="false" form:apply-design-mode="false" /><table:table-column table:style-name="co1" table:default-cell-style-name="ce15" table:number-columns-repeated="256" /><table:table-row table:style-name="ro1"><table:table-cell office:value-type="float" office:value="37"><text:p>37</text:p></table:table-cell><table:table-cell office:value-type="string"><text:p>26</text:p></table:table-cell><table:table-cell table:number-columns-repeated="1022" /></table:table-row></table:table><table:table table:name="L'autre feuille" table:style-name="ta2" table:print="false"><office:forms form:automatic-focus="false" form:apply-design-mode="false" /><table:table-column table:style-name="co2" table:default-cell-style-name="ce15" table:number-columns-repeated="256" /><table:table-row table:style-name="ro1"><table:table-cell office:value-type="string"><text:p>1</text:p></table:table-cell><table:table-cell office:value-type="string"><text:p>2</text:p></table:table-cell><table:table-cell office:value-type="string"><text:p>3</text:p></table:table-cell><table:table-cell office:value-type="string"><text:p>4</text:p></table:table-cell><table:table-cell office:value-type="string"><text:p>5</text:p></table:table-cell><table:table-cell table:number-columns-repeated="1019" /></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>`
    );
    /*
    const root = doc.documentElement;
    setNamespaces(root);
    
    const body = doc.createElement('office:body');
    const spreadsheet = doc.createElement('office:spreadsheet');
    
    for (const [sheetName, sheetContent] of sheetsData) {
        const table = createTableElement(doc, sheetName, sheetContent);
        spreadsheet.appendChild(table);
    }
    
    body.appendChild(spreadsheet);
    root.appendChild(body);*/
    
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
    table.setAttribute('table:style-name', "ta1"); // PPP hardcoded
    
    //<table:table-column table:style-name="co1" table:default-cell-style-name="ce15" table:number-columns-repeated="256" />
    const tableColumn = doc.createElement('table:table-column');
    table.setAttribute('table:style-name', "co1"); // PPP hardcoded
    table.setAttribute('table:number-columns-repeated', String(Math.max(...sheetContent.map(row => row.length)))); // PPP hardcoded

    sheetContent.forEach(rowContent => {
        const row = createRowElement(doc, rowContent);
        table.appendChild(row);
    });
    
    return table;
}

/**
 * Crée l'élément XML pour une ligne de la feuille de calcul
 * @param {Document} doc 
 * @param {SheetRowRawContent} rowContent
 * @returns {Element}
 */
function createRowElement(doc, rowContent) {
    const row = doc.createElement('table:table-row');
    row.setAttribute('table:style-name', "ro1"); // PPP hardcoded
    
    rowContent.forEach(cell => {
        const cellElement = createCellElement(doc, cell);
        row.appendChild(cellElement);
    });
    
    return row;
}

/**
 * @param {any} _styles 
 * @param {(str: string) => Document} parseXML - Function to parse XML content.
 * @param {(doc: Document) => string} serializeXML - Function to parse XML content.
 * @returns {string}
 */
function createStylesXml(_styles, parseXML, serializeXML) {
    // adapted from https://git.sheetjs.com/sheetjs/sheetjs/src/commit/235ed7ccfb9fed4aaadce3a2d693027c72d314ae/bits/81_writeods.js#L2
    const doc = parseXML(
        `<?xml version="1.0" encoding="utf-8"?><office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0" xmlns:math="http://www.w3.org/1998/Math/MathML" xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0" xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0" xmlns:ooo="http://openoffice.org/2004/office" xmlns:ooow="http://openoffice.org/2004/writer" xmlns:oooc="http://openoffice.org/2004/calc" xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2" xmlns:dom="http://www.w3.org/2001/xml-events" office:version="1.2"><office:font-face-decls><style:font-face  style:name="Arial" svg:font-family="Arial" /></office:font-face-decls><office:styles><number:number-style style:name="N0"><number:number number:min-integer-digits="1" /></number:number-style><number:percentage-style style:name="N9"><number:number number:decimal-places="0" number:min-integer-digits="1" /><number:text>%</number:text></number:percentage-style><number:number-style style:name="N41P0" style:volatile="true"><number:text>  </number:text><number:number number:decimal-places="0" number:min-integer-digits="1" number:grouping="true" /><number:text> </number:text></number:number-style><number:number-style style:name="N41P1" style:volatile="true"><number:text>  (</number:text><number:number number:decimal-places="0" number:min-integer-digits="1" number:grouping="true" /><number:text>)</number:text></number:number-style><number:text-style style:name="N41P2" style:volatile="true"><number:text>  - </number:text></number:text-style><number:text-style style:name="N41"><number:text> </number:text><number:text-content /><number:text> </number:text><style:map style:condition="value()&gt;0" style:apply-style-name="N41P0" /><style:map style:condition="value()&lt;0" style:apply-style-name="N41P1" /><style:map style:condition="value()=0" style:apply-style-name="N41P2" /></number:text-style><number:number-style style:name="N42P0" style:volatile="true"><number:text> ¤ </number:text><number:number number:decimal-places="0" number:min-integer-digits="1" number:grouping="true" /><number:text> </number:text></number:number-style><number:number-style style:name="N42P1" style:volatile="true"><number:text> ¤ (</number:text><number:number number:decimal-places="0" number:min-integer-digits="1" number:grouping="true" /><number:text>)</number:text></number:number-style><number:text-style style:name="N42P2" style:volatile="true"><number:text> ¤ - </number:text></number:text-style><number:text-style style:name="N42"><number:text> </number:text><number:text-content /><number:text> </number:text><style:map style:condition="value()&gt;0" style:apply-style-name="N42P0" /><style:map style:condition="value()&lt;0" style:apply-style-name="N42P1" /><style:map style:condition="value()=0" style:apply-style-name="N42P2" /></number:text-style><number:number-style style:name="N43P0" style:volatile="true"><number:text>  </number:text><number:number number:min-integer-digits="1" number:grouping="true" number:decimal-places="2" /><number:text> </number:text></number:number-style><number:number-style style:name="N43P1" style:volatile="true"><number:text>  (</number:text><number:number number:min-integer-digits="1" number:grouping="true" number:decimal-places="2" /><number:text>)</number:text></number:number-style><number:text-style style:name="N43P2" style:volatile="true"><number:text>  - </number:text></number:text-style><number:text-style style:name="N43"><number:text> </number:text><number:text-content /><number:text> </number:text><style:map style:condition="value()&gt;0" style:apply-style-name="N43P0" /><style:map style:condition="value()&lt;0" style:apply-style-name="N43P1" /><style:map style:condition="value()=0" style:apply-style-name="N43P2" /></number:text-style><number:number-style style:name="N44P0" style:volatile="true"><number:text> ¤ </number:text><number:number number:min-integer-digits="1" number:grouping="true" number:decimal-places="2" /><number:text> </number:text></number:number-style><number:number-style style:name="N44P1" style:volatile="true"><number:text> ¤ (</number:text><number:number number:min-integer-digits="1" number:grouping="true" number:decimal-places="2" /><number:text>)</number:text></number:number-style><number:text-style style:name="N44P2" style:volatile="true"><number:text> ¤ - </number:text></number:text-style><number:text-style style:name="N44"><number:text> </number:text><number:text-content /><number:text> </number:text><style:map style:condition="value()&gt;0" style:apply-style-name="N44P0" /><style:map style:condition="value()&lt;0" style:apply-style-name="N44P1" /><style:map style:condition="value()=0" style:apply-style-name="N44P2" /></number:text-style><style:style style:name="Default" style:family="table-cell" style:data-style-name="N0"><style:table-cell-properties fo:border="none" style:vertical-align="middle" fo:background-color="transparent" style:cell-protect="protected" /><style:paragraph-properties /><style:text-properties  style:font-name="Arial" style:font-name-asian="Arial" fo:font-weight="normal" style:font-weight-asian="normal" style:font-weight-complex="normal" fo:font-style="normal" style:font-style-asian="normal" style:font-style-complex="normal" style:text-line-through-style="none" fo:font-size="10pt" style:font-size-asian="10pt" style:font-size-complex="10pt" fo:color="#000000" style:text-underline-style="none" /></style:style><style:style style:name="Percent" style:family="table-cell" style:data-style-name="N9" /><style:style style:name="Currency" style:family="table-cell" style:data-style-name="N44" /><style:style style:name="Currency [0]" style:family="table-cell" style:data-style-name="N42" /><style:style style:name="Comma" style:family="table-cell" style:data-style-name="N43" /><style:style style:name="Comma [0]" style:family="table-cell" style:data-style-name="N41" /></office:styles><office:automatic-styles><style:page-layout style:name="pm1"><style:page-layout-properties style:print-orientation="portrait" fo:page-width="8.5in" fo:page-height="11in" style:scale-to="100%" style:print-page-order="ttb" fo:margin-left="0.75in" fo:margin-right="0.75in" fo:margin-top="0.5in" fo:margin-bottom="0.5in" fo:background-color="transparent" style:first-page-number="continue" /><style:header-style><style:header-footer-properties fo:min-height="0.5in" fo:margin-left="0in" fo:margin-right="0in" fo:margin-bottom="0in" /></style:header-style><style:footer-style><style:header-footer-properties fo:min-height="0.5in" fo:margin-left="0in" fo:margin-right="0in" fo:margin-top="0in" /></style:footer-style></style:page-layout><style:page-layout style:name="pm2"><style:page-layout-properties style:print-orientation="portrait" fo:page-width="8.5in" fo:page-height="11in" style:scale-to="100%" style:print-page-order="ttb" fo:margin-left="0.75in" fo:margin-right="0.75in" fo:margin-top="0.5in" fo:margin-bottom="0.5in" fo:background-color="transparent" style:first-page-number="continue" /><style:header-style><style:header-footer-properties fo:min-height="0.5in" fo:margin-left="0in" fo:margin-right="0in" fo:margin-bottom="0in" /></style:header-style><style:footer-style><style:header-footer-properties fo:min-height="0.5in" fo:margin-left="0in" fo:margin-right="0in" fo:margin-top="0in" /></style:footer-style></style:page-layout></office:automatic-styles><office:master-styles><style:master-page style:name="PageStyle_5f_La feuille" style:display-name="PageStyle_La feuille" style:page-layout-name="pm1"><style:header style:display="false" /><style:header-left style:display="false" /><style:footer style:display="false" /><style:footer-left style:display="false" /></style:master-page><style:master-page style:name="PageStyle_5f_L'autre feuille" style:display-name="PageStyle_L'autre feuille" style:page-layout-name="pm2"><style:header style:display="false" /><style:header-left style:display="false" /><style:footer style:display="false" /><style:footer-left style:display="false" /></style:master-page></office:master-styles></office:document-styles>`
    );

    /*const root = doc.documentElement;
    setNamespaces(root);*/
    
    // PPP add styles
    
    return serializeXML(doc);
}

/**
 * @param {any} _metadata 
 * @param {(str: string) => Document} parseXML - Function to parse XML content.
 * @param {(doc: Document) => string} serializeXML - Function to parse XML content.
 * @returns {string}
 */
function createMetaXml(_metadata, parseXML, serializeXML) {
    const doc = parseXML(
        `<?xml version="1.0" encoding="utf-8"?><office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:ooo="http://openoffice.org/2004/office" office:version="1.2"><office:meta><meta:generator>https://github.com/DavidBruant/ods-xlsx</meta:generator><meta:creation-date>2024-10-23T10:34:24Z</meta:creation-date></office:meta></office:document-meta>`
    );
    
    /*const root = doc.documentElement;
    setNamespaces(root);*/

    // PPP add actual metadata
    
    return serializeXML(doc);
}

/**
 * @param {any} _metadata 
 * @param {(str: string) => Document} parseXML - Function to parse XML content.
 * @param {(doc: Document) => string} serializeXML - Function to parse XML content.
 * @returns {string}
 */
function createSettingsXml(_metadata, parseXML, serializeXML) {
    const doc = parseXML(
        `<?xml version="1.0" encoding="utf-8"?><office:document-settings xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0" xmlns:ooo="http://openoffice.org/2004/office" office:version="1.2"><office:settings><config:config-item-set config:name="ooo:view-settings"><config:config-item-map-indexed config:name="Views"><config:config-item-map-entry><config:config-item config:name="ViewId" config:type="string">View1</config:config-item><config:config-item-map-named config:name="Tables"><config:config-item-map-entry config:name="La feuille"><config:config-item config:name="PositionLeft" config:type="int">0</config:config-item><config:config-item config:name="PositionTop" config:type="int">0</config:config-item><config:config-item config:name="PositionRight" config:type="int">0</config:config-item><config:config-item config:name="PositionBottom" config:type="int">0</config:config-item><config:config-item config:name="ZoomType" config:type="short">0</config:config-item><config:config-item config:name="ZoomValue" config:type="int">100</config:config-item><config:config-item config:name="PageViewZoomValue" config:type="int">100</config:config-item><config:config-item config:name="ShowGrid" config:type="boolean">true</config:config-item></config:config-item-map-entry><config:config-item-map-entry config:name="L'autre feuille"><config:config-item config:name="PositionLeft" config:type="int">0</config:config-item><config:config-item config:name="PositionTop" config:type="int">0</config:config-item><config:config-item config:name="PositionRight" config:type="int">0</config:config-item><config:config-item config:name="PositionBottom" config:type="int">0</config:config-item><config:config-item config:name="ZoomType" config:type="short">0</config:config-item><config:config-item config:name="ZoomValue" config:type="int">100</config:config-item><config:config-item config:name="PageViewZoomValue" config:type="int">100</config:config-item><config:config-item config:name="ShowGrid" config:type="boolean">true</config:config-item></config:config-item-map-entry></config:config-item-map-named><config:config-item config:name="ActiveTable" config:type="string">La feuille</config:config-item><config:config-item config:name="ShowPageBreakPreview" config:type="boolean">false</config:config-item><config:config-item config:name="ShowZeroValues" config:type="boolean">true</config:config-item><config:config-item config:name="HasColumnRowHeaders" config:type="boolean">true</config:config-item><config:config-item config:name="ShowGrid" config:type="boolean">true</config:config-item><config:config-item config:name="GridColor" config:type="long">12632256</config:config-item><config:config-item config:name="HasSheetTabs" config:type="boolean">true</config:config-item><config:config-item config:name="HorizontalScrollbarWidth" config:type="int">600</config:config-item></config:config-item-map-entry></config:config-item-map-indexed></config:config-item-set></office:settings></office:document-settings>`
    );
    
    /*const root = doc.documentElement;
    setNamespaces(root);*/
    
    return serializeXML(doc);
}




/**
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

    root.setAttribute('office:version', officeVersion);
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