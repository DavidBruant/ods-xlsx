import test from 'ava';
import {join} from 'node:path';

import {fillOdtTemplate, getOdtTemplate, getOdtTextContent} from '../scripts/fillOdtTemplate.js'


test('basic template filling with variable substitution', async t => {
    

    const templatePath = join(import.meta.dirname, './data/template-anniversaire.odt')
    const templateContent = `Yo {nom} ! 
Tu es né.e le {dateNaissance}

Bonjoir ☀️
`

	const data = {
        nom: 'David Bruant',
        dateNaissance: '8 mars 1987'
    }

    const odtTemplate = await getOdtTemplate(templatePath)
    const templateTextContent = await getOdtTextContent(odtTemplate)
    t.deepEqual(templateTextContent, templateContent)

    const odtResult = await fillOdtTemplate(odtTemplate, data)

    const odtResultTextContent = await getOdtTextContent(odtResult)
    t.deepEqual(odtResultTextContent, `Yo David Bruant ! 
Tu es né.e le 8 mars 1987

Bonjoir ☀️
`)

});



test('basic template filling with {#each}', async t => {
    const templatePath = join(import.meta.dirname, './data/enum-courses.odt')
    const templateContent = `🧺 La liste de courses incroyable 🧺

{#each listeCourses as élément}
{élément}
{/each}
`

	const data = {
        listeCourses : [
            'Radis',
            `Jus d'orange`,
            'Pâtes à lasagne (fraîches !)'
        ]
    }

    const odtTemplate = await getOdtTemplate(templatePath)

    const templateTextContent = await getOdtTextContent(odtTemplate)

    t.deepEqual(templateTextContent, templateContent)

    const odtResult = await fillOdtTemplate(odtTemplate, data)

    const odtResultTextContent = await getOdtTextContent(odtResult)
    t.deepEqual(odtResultTextContent, `🧺 La liste de courses incroyable 🧺

Radis
Jus d'orange
Pâtes à lasagne (fraîches !)
`)


});



test('basic template filling with {#each} generating a list', async t => {
    const templatePath = join(import.meta.dirname, './data/liste-courses.odt')
    const templateContent = `🧺 La liste de courses incroyable 🧺

- {#each listeCourses as élément}
- {élément}
- {/each}
`

	const data = {
        listeCourses : [
            'Radis',
            `Jus d'orange`,
            'Pâtes à lasagne (fraîches !)'
        ]
    }

    const odtTemplate = await getOdtTemplate(templatePath)

    const templateTextContent = await getOdtTextContent(odtTemplate)

    t.deepEqual(templateTextContent, templateContent)

    const odtResult = await fillOdtTemplate(odtTemplate, data)

    const odtResultTextContent = await getOdtTextContent(odtResult)
    t.deepEqual(odtResultTextContent, `🧺 La liste de courses incroyable 🧺

- Radis
- Jus d'orange
- Pâtes à lasagne (fraîches !)
`)


});
