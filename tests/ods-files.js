import {readFile} from 'node:fs/promises'

import test from 'ava';

import {getODSTableRawContent} from '../scripts/main.js'

test('.ods file with table:number-columns-repeated attribute in cell', async t => {
    const repeatedCellFileContent = (await readFile('./tests/data/cellules-répétées.ods')).buffer

	const table = await getODSTableRawContent(repeatedCellFileContent);

    const feuille1 = table.get('Feuille 1')

    console.log('yo repeated', feuille1)
    t.deepEqual(feuille1[0].length, feuille1[1].length, `First and second row should have the same number of columns`)
});



test('.ods cells with dates should be recognized', async t => {
    const odsFileWithDates = (await readFile('./tests/data/cellules avec dates.ods')).buffer
	const table = await getODSTableRawContent(odsFileWithDates);

    const feuille1 = table.get('Feuille1')

    console.log('yo dates', feuille1)

    const row1 = feuille1[0]
    t.deepEqual(row1[0].value, 'Nom')
    t.deepEqual(row1[1].value, 'Date de naissance')

    const row2 = feuille1[1]
    t.deepEqual(row2[0].value, 'Dav')
    t.deepEqual(row2[1].type, 'date')
    t.deepEqual(row2[1].value, '1987-03-08')

    const row3 = feuille1[2]
    t.deepEqual(row3[0].value, 'Fanny')
    t.deepEqual(row3[1].type, 'date')
    t.deepEqual(row3[1].value, '1986-06-01')
});