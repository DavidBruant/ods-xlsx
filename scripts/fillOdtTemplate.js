import {readFile} from 'node:fs/promises'

import {ZipReader, Uint8ArrayReader, TextWriter} from '@zip.js/zip.js';
import {DOMParser} from '@xmldom/xmldom';


// fillOdtTemplate, getOdtTemplate, getOdtTextContent

/** @typedef {ArrayBuffer} ODTFile */

/**
 * 
 * @param {string} path 
 * @returns {Promise<ODTFile>}
 */
export async function getOdtTemplate(path){
    const buffer = await readFile(path)
    return buffer.buffer
}


/**
 * Extracts plain text content from an ODT file, preserving line breaks
 * @param {ArrayBuffer} odtFile - The ODT file as an ArrayBuffer
 * @returns {Promise<string>} Extracted text content
 */
export async function getOdtTextContent(odtFile) {
    try {
        // Create a reader from the ArrayBuffer
        const reader = new ZipReader(new Uint8ArrayReader(new Uint8Array(odtFile)));
        
        // Get all entries in the zip file
        const entries = await reader.getEntries();
        
        // Find the content.xml file (where text is stored in ODT)
        const contentEntry = entries.find(entry => entry.filename === 'content.xml');
        
        if (!contentEntry) {
            throw new Error('No content.xml found in the ODT file');
        }
        
        // Extract the content.xml as text
        const contentText = await contentEntry.getData(new TextWriter());
        
        // Parse the XML to extract plain text
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(contentText, 'text/xml');
        
        // Extract text from text:p elements, preserving paragraphs
        const paragraphs = xmlDoc.getElementsByTagName('text:p');
        const textLines = Array.from(paragraphs).map(p => p.textContent || '');
        
        // Close the zip reader
        await reader.close();
        
        // Join paragraphs with newlines to preserve structure
        return textLines.join('\n');
    } catch (error) {
        console.error('Error extracting ODT content:', error);
        throw error;
    }
}

export function fillOdtTemplate(){
    throw `PPP`
}