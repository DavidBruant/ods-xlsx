import {join} from 'node:path';

import {fillOdtTemplate, getOdtTemplate} from '../scripts/fillOdtTemplate.js'

/*
const templatePath = join(import.meta.dirname, '../tests/data/template-anniversaire.odt')
const data = {
    nom: 'David Bruant',
    dateNaissance: '8 mars 1987'
}
*/


/*
const templatePath = join(import.meta.dirname, '../tests/data/liste-courses.odt')
const data = {
    listeCourses : [
        'Radis',
        `Jus d'orange`,
        'Pâtes à lasagne (fraîches !)'
    ]
}
*/


const templatePath = join(import.meta.dirname, '../tests/data/liste-fruits-et-légumes.odt')
const data = {
    fruits : [
        'Pastèque 🍉',
        `Kiwi 🥝`,
        'Banane 🍌'
    ],
    légumes: [
        'Champignon 🍄‍🟫',
        'Avocat 🥑',
        'Poivron 🫑'
    ]
}


const odtTemplate = await getOdtTemplate(templatePath)
const odtResult = await fillOdtTemplate(odtTemplate, data)

process.stdout.write(new Uint8Array(odtResult))