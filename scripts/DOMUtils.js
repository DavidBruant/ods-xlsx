import { Node } from '@xmldom/xmldom';


/*
    Since we're using xmldom in Node.js context, the entire DOM API is not implemented
    Functions here are helpers whild xmldom becomes more complete
*/

/**
 * Traverses a DOM tree starting from the given element and applies the visit function
 * to each Element node encountered in tree order (depth-first).
 * 
 * @param {Element} element - The starting DOM Element for traversal
 * @param {Function} visit - Function to be called on each Element, receives the Element as its argument
 */
export function traverse(element, visit) {
    const children = Array.from(element.childNodes)
        .filter(child => child.nodeType === Node.ELEMENT_NODE);

    for (const child of children) {
        // @ts-ignore
        traverse(child, visit);
    }

    visit(element);
}
