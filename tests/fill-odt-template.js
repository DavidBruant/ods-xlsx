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
    t.deepEqual(templateTextContent, templateContent, 'reconnaissance du template')

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

    t.deepEqual(templateTextContent, templateContent, 'reconnaissance du template')

    const odtResult = await fillOdtTemplate(odtTemplate, data)

    const odtResultTextContent = await getOdtTextContent(odtResult)
    t.deepEqual(odtResultTextContent, `🧺 La liste de courses incroyable 🧺

Radis
Jus d'orange
Pâtes à lasagne (fraîches !)
`)


});



test('template filling with {#each} generating a list', async t => {
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

    t.deepEqual(templateTextContent, templateContent, 'reconnaissance du template')

    const odtResult = await fillOdtTemplate(odtTemplate, data)

    const odtResultTextContent = await getOdtTextContent(odtResult)
    t.deepEqual(odtResultTextContent, `🧺 La liste de courses incroyable 🧺

- Radis
- Jus d'orange
- Pâtes à lasagne (fraîches !)
`)


});


test('template filling with 2 sequential {#each}', async t => {
    const templatePath = join(import.meta.dirname, './data/liste-fruits-et-légumes.odt')
    const templateContent = `Liste de fruits et légumes

Fruits
{#each fruits as fruit}
{fruit}
{/each}

Légumes
{#each légumes as légume}
{légume}
{/each}
`

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

    const templateTextContent = await getOdtTextContent(odtTemplate)    
    t.deepEqual(templateTextContent, templateContent, 'reconnaissance du template')

    const odtResult = await fillOdtTemplate(odtTemplate, data)

    const odtResultTextContent = await getOdtTextContent(odtResult)
    t.deepEqual(odtResultTextContent, `Liste de fruits et légumes

Fruits
Pastèque 🍉
Kiwi 🥝
Banane 🍌

Légumes
Champignon 🍄‍🟫
Avocat 🥑
Poivron 🫑
`)

});



test('template filling with nested {#each}s', async t => {
    const templatePath = join(import.meta.dirname, './data/légumes-de-saison.odt')
    const templateContent = `Légumes de saison

{#each légumesSaison as saisonLégumes}
{saisonLégumes.saison}
- {#each saisonLégumes.légumes as légume}
- {légume}
- {/each}

{/each}
`

	const data = {
        légumesSaison : [
            {
                saison: 'Printemps',
                légumes: [
                    'Asperge',
                    'Betterave',
                    'Blette'
                ]
            },
            {
                saison: 'Été',
                légumes: [
                    'Courgette',
                    'Poivron',
                    'Laitue'
                ]
            },
            {
                saison: 'Automne',
                légumes: [
                    'Poireau',
                    'Potiron',
                    'Brocoli'
                ]
            },
            {
                saison: 'Hiver',
                légumes: [
                    'Radis',
                    'Chou de Bruxelles',
                    'Frisée'
                ]
            }
        ]
    }

    const odtTemplate = await getOdtTemplate(templatePath)

    const templateTextContent = await getOdtTextContent(odtTemplate)    
    t.deepEqual(templateTextContent, templateContent, 'reconnaissance du template')

    const odtResult = await fillOdtTemplate(odtTemplate, data)

    const odtResultTextContent = await getOdtTextContent(odtResult)
    t.deepEqual(odtResultTextContent, `Légumes de saison

Printemps
- Asperge
- Betterave
- Blette

Été
- Courgette
- Poivron
- Laitue

Automne
- Poireau
- Potiron
- Brocoli

Hiver
- Radis
- Chou de Bruxelles
- Frisée

`)

});