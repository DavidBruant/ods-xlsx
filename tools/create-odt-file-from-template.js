import {join} from 'node:path';

import {fillOdtTemplate, getOdtTemplate} from '../scripts/fillOdtTemplate.js'

const templatePath = join(import.meta.dirname, '../tests/data/template-anniversaire.odt')
const odtTemplate = await getOdtTemplate(templatePath)

const data = {
    nom: 'David Bruant',
    dateNaissance: '8 mars 1987'
}

const odtResult = await fillOdtTemplate(odtTemplate, data)

process.stdout.write(new Uint8Array(odtResult))