import test from 'ava';
import {join} from 'node:path';

import {fillOdtTemplate, getOdtTemplate, getOdtTextContent} from '../scripts/fillOdtTemplate.js'

const templatePath = join(import.meta.dirname, './data/template-anniversaire.odt')
const templateContent = `Yo {nom} ! 
Tu es né.e le {dateNaissance}

Bonjoir ☀️`


test('basic template filling', async t => {
	const data = {
        nom: 'David Bruant',
        dateNaissance: '8 mars 1987'
    }

    const odtTemplate = await getOdtTemplate(templatePath)

    const templateTextContent = await getOdtTextContent(odtTemplate)
    console.log('templateTextContent', templateTextContent)
    t.deepEqual(templateTextContent, templateContent)

    const odtResult = await fillOdtTemplate(odtTemplate, data)

    const odtResultTextContent = await getOdtTextContent(odtTemplate)
    t.deepEqual(templateTextContent, `Yo David Bruant ! 
Tu es né.e le 8 mars 1987

Bonjoir ☀️`)

});
