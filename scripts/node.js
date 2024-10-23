//@ts-check

import {DOMParser, XMLSerializer} from '@xmldom/xmldom'

import {
    _getODSTableRawContent, 
    _getXLSXTableRawContent
} from './shared.js'

import {_createOdsFile} from './createOdsFile.js'

/** @import {SheetName, SheetRawContent} from './types.js' */


function parseXML(str){
    return (new DOMParser()).parseFromString(str, 'application/xml');
}

function serializeXML(doc){
    return (new XMLSerializer()).serializeToString(doc);
}

/**
 * @param {ArrayBuffer} odsArrBuff
 * @returns {ReturnType<_getODSTableRawContent>}
 */
export function getODSTableRawContent(odsArrBuff){
    return _getODSTableRawContent(odsArrBuff, parseXML)
}

/**
 * @param {ArrayBuffer} xlsxArrBuff
 * @returns {ReturnType<_getXLSXTableRawContent>}
 */
export function getXLSXTableRawContent(xlsxArrBuff){
    return _getXLSXTableRawContent(xlsxArrBuff, parseXML)
}

/**
 * Crée un fichier .ods à partir d'un Map de feuilles de calcul
 * @param {Map<SheetName, SheetRawContent>} sheetsData
 * @returns {Promise<Uint8Array>}
 */
export function createOdsFile(sheetsData){
    return _createOdsFile(sheetsData, parseXML, serializeXML)
}


export {
    // table-level exports
    tableWithoutEmptyRows,
    tableRawContentToValues,
    tableRawContentToStrings,
    tableRawContentToObjects, 

    // sheet-level exports
    sheetRawContentToObjects,
    sheetRawContentToStrings,

    // row-level exports
    rowRawContentToStrings,
    isRowNotEmpty,

    // cell-level exports
    cellRawContentToStrings,
    convertCellValue
} from './shared.js'

