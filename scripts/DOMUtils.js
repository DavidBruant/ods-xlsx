import { Node } from '@xmldom/xmldom';


/*
    Since we're using xmldom in Node.js context, the entire DOM API is not implemented
    Functions here are helpers whild xmldom becomes more complete
*/

/**
 * Traverses a DOM tree starting from the given element and applies the visit function
 * to each Element node encountered in tree order (depth-first).
 * 
 * @param {Node} node
 * @param {(n : Node) => void} visit
 */
export function traverse(node, visit) {
    //console.log('traverse', node.nodeType, node.nodeName)
    
    for (const child of Array.from(node.childNodes)) {
        traverse(child, visit);
    }

    visit(node);
}
