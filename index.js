const xlsx = require('xlsx');
const _ = require('lodash');

// 为便于处理国家产品名全部小写
const kNationBrands = {
    france: ['lenovo', 'asus', 'lactel', 'candia'],
    china: ['lenovo', 'asus', 'brightdairy', 'yili'],
};
// 产品类别，如 asus 是第1类产品的第2号品牌
const kBrandCategory = {lenovo: [1, 1], asus: [1, 2], lactel: [2, 1], candia: [2, 2], brightdairy: [2, 1], yili:[2, 2] };
// 每个问题的编号区间（前闭后开），如 _.range(1,3) 代表 [1, 2]
const kBrandRelativeCategory = [
    { purchase: [] },
    { bc: _.range(1, 4) }, 
    { cr: _.range(1, 7) },
    { pq: _.range(1, 3) },
    { ics: _.range(1, 3) },
    { lpr: _.range(1, 4) },
    { pr: _.range(1, 3) },
    { pcp: _.range(1, 5) },
];
const kBrandIrrelativeCategory = [
    { ua: _.range(1, 4) },
    { ci: _.range(1, 7) },
    { pd: _.range(1, 5) },
    { gender: [] },
    { age: [] },
    { education: [] },
    { occupation: [] },
    { salary: [] },
];

const kBrandQuestionUpper = 24;
const kCultureSocioQuestionUpper = 42;
const kBrandQuestionRange = _.range(1, kBrandQuestionUpper);
const kCultureSocioQuestionRange = _.range(24, kCultureSocioQuestionUpper);

const kColumnHeaderFormat = /^(\d+)\.\s*(.+)$/;

function solve(nation) {
    const wbChina = xlsx.readFile(`${nation}.xlsx`);
    const shChina = wbChina.Sheets[wbChina.SheetNames[0]];
    const jsonChina = xlsx.utils.sheet_to_json(shChina);

    const brands = kNationBrands[nation];
    const respondents = [];
    for (const row of jsonChina) {
        const brandRate = _.zipObject(brands, brands.map(() => [undefined]));
        const cultureSocioRate = [];
        for (const [key, value] of Object.entries(row)) {
            const match = key.toLowerCase().match(kColumnHeaderFormat);
            if (!match) continue;
            let [, no, title] = match;
            no = parseInt(no);
            if (_.inRange(no, kBrandQuestionRange[0], kBrandQuestionUpper)) {
                brandRate[title.toLowerCase()][no] = value;
            } else if (_.inRange(no, kCultureSocioQuestionRange[0], kCultureSocioQuestionUpper)) {
                cultureSocioRate[no - kBrandQuestionUpper] = value;
            } else {
                // 忽略 Which would you choose 问题
                continue;
            }
        }
        respondents.push({ brandRate, cultureSocioRate });
    }
    
    const outputHeader = ['No.', 'category', 'branch'];
    for (const obj of [...kBrandRelativeCategory, ...kBrandIrrelativeCategory]) {
        const k = Object.keys(obj)[0];
        const vs = Object.values(obj)[0];
        if (vs.length > 0) {
            const arr = vs.map(v => k + v);
            outputHeader.push(...arr);
        } else {
            outputHeader.push(k);
        }
    }
    
    const outputData = [outputHeader];
    for (let i = 0; i < respondents.length; i++) {
        const respondent = respondents[i];
        for (let brand of brands) {
            const brandData = [i + 1];
            brandData.push(...kBrandCategory[brand], ...respondent.brandRate[brand].slice(1), ...respondent.cultureSocioRate);
            outputData.push(brandData);
        }
    }
    
    const filename = `output_${nation}.xlsx`;
    const ws_name = "Sheet1";
    const wb = xlsx.utils.book_new();
    const ws = xlsx.utils.aoa_to_sheet(outputData);
    /* add worksheet to workbook */
    xlsx.utils.book_append_sheet(wb, ws, ws_name);
    /* write workbook */
    xlsx.writeFile(wb, filename);
}

solve('china');
solve('france');