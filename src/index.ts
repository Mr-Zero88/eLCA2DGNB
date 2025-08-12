import jsdom from "jsdom";
import exeljs from "exceljs";
import fs from "fs/promises";

const sid = "ee2141868f235d36e8903e76629fa30a";
const projectId = "68038";

(async () => {
    const filePath = process.argv[2];
    if (!filePath) throw new Error("Please provide the path to the Excel file as an argument.");
    const data = await getELCAData(sid, projectId);
    const placeholders = generatePlaceholders(data);
    await editExcelFile(filePath, placeholders);
    console.log("Done.");
})();

interface IndicatorData {
    unit: string;
    manufacture: number;
    disposal: number;
    servicing: number;
    total: number;
    potential: number;
}

interface CategoryData {
    [indicatorName: string]: IndicatorData;
}

interface ELCAData {
    [categoryName: string]: CategoryData;
}

async function getELCAData(sid: string, projectId: string): Promise<ELCAData> {
    const response = await fetch("https://www.bauteileditor.de/project-reports/summaryElementTypes/?_isBaseReq=true", {
        "headers": {
            "accept": "*/*",
            "accept-language": "en-US,en;q=0.9",
            "priority": "u=1, i",
            "sec-ch-ua": "\"Not/A)Brand\";v=\"8\", \"Chromium\";v=\"132\", \"Google Chrome\";v=\"132\"",
            "sec-ch-ua-mobile": "?0",
            "sec-ch-ua-platform": "\"Windows\"",
            "sec-fetch-dest": "empty",
            "sec-fetch-mode": "cors",
            "sec-fetch-site": "same-origin",
            "x-hash-url": "/project-reports/summaryElementTypes/",
            "x-requested-with": "XMLHttpRequest",
            "cookie": `sid=${sid}`,
            "Referer": `https://www.bauteileditor.de/projects/${projectId}/`
        },
        "body": null,
        "method": "GET"
    });
    const data = await response.json();
    const reportElement = data["Elca\\View\\Report\\ElcaReportElementTypeEffectsView"];
    if (!reportElement) {
        throw new Error("No ELCA data found: " + Object.keys(data).join(", "));
    }

    let dom = new jsdom.JSDOM(reportElement);
    let root = dom.window.document.documentElement;
    let categories: ELCAData = {};
    Array.from(root.querySelectorAll(".print-content > .category > li")).forEach((categoryElement) => {
        let categoryName = parseCategoryName(categoryElement.querySelector(":scope > h1")?.textContent?.trim() || "NULL")
        let categoryData: CategoryData = {};
        Array.from(categoryElement.querySelectorAll(":scope > table > tbody > tr")).forEach((valueElement) => {
            let dataElements = Array.from(valueElement.querySelectorAll(":scope > td"));
            let indicatorName = parseIndicatorName(dataElements[0]?.textContent?.trim() || "NULL");
            let indicatorData: IndicatorData = {
                unit: parseUnit(dataElements[1]?.textContent?.trim() || "NULL"),
                manufacture: Number(dataElements[2]?.textContent?.trim() || "0"),
                disposal: Number(dataElements[3]?.textContent?.trim() || "0"),
                servicing: Number(dataElements[4]?.textContent?.trim() || "0"),
                total: Number(dataElements[5]?.textContent?.trim() || "0"),
                potential: Number(dataElements[6]?.textContent?.trim() || "0"),
            };
            if (!indicatorName) {
                throw new Error("No indicator name found in element: " + valueElement.outerHTML);
            }
            categoryData[indicatorName] = indicatorData;
        });
        categories[categoryName] = categoryData;
    });

    return categories;
}

function parseCategoryName(categoryName: string): string {
    if (categoryName == "Total / Construction") {
        return "TOTAL";
    }
    if (categoryName.match(/^\d/)) {
        return `KG${categoryName.split(' ')[0]}`;
    }
    throw new Error(`Unknown category name: ${categoryName}`);
}

function parseIndicatorName(indicatorName: string): string {
    switch (indicatorName) {
        case "GWP":
        case "ODP":
        case "POCP":
        case "AP":
        case "EP":
        case "PENRT":
        case "PENRM":
        case "PENRE":
        case "PERT":
        case "PERM":
        case "PERE":
        case "SM":
        case "FW":
            return indicatorName;
        case "Total PE":
            return "TPE";
        case "ADP elem.":
            return "ADPE";
        case "ADP fossil":
            return "ADPF";
        default:
            throw new Error(`Unknown indicator name: ${indicatorName}`);
    }
}

function parseUnit(unit: string): string {
    switch (unit) {
        case "kg CO2 equiv.":
            return "CO2";
        case "kg R11 equiv.":
            return "CFC11";
        case "kg ethene equiv.":
            return "C2H4";
        case "kg SO2 eqv.":
            return "SO2";
        case "kg PO4 equiv.":
            return "PO4";
        case "kg Sb equiv.":
            return "SB";
        case "MJ":
            return "MJ";
        case "kg":
            return "KG";
        case "m3":
            return "M3";
        default:
            throw new Error(`Unknown unit: ${unit}`);
    }
}

function generatePlaceholders(data: ELCAData): { [key: string]: number } {
    let placeholders: { [key: string]: number } = {};

    for (const category in data) {
        for (const indicator in data[category]) {
            let { unit, ...indicatorData } = data[category][indicator];
            for (const key in indicatorData) {
                let placeholderKey = `${category}/${indicator}/${unit}/${key.toUpperCase()}`;
                if (placeholders[placeholderKey] !== undefined)
                    throw new Error(`Duplicate placeholder key: ${placeholderKey}`);
                placeholders[placeholderKey] = indicatorData[key];
            }
        }
    }

    return placeholders;
}

async function editExcelFile(filePath: string, placeholders: { [key: string]: number }) {
    const workbook = new exeljs.Workbook();
    await workbook.xlsx.readFile(filePath);
    let version = await getFileVersion(workbook);
    console.log(`Using template version: ${version}`);
    let placeholderPlacement = await getTemplatePlaceholderPlacement(version);
    replacePlaceholdersInExcel(workbook, placeholders, placeholderPlacement);
    await workbook.xlsx.writeFile(filePath);
}

async function replacePlaceholdersInExcel(
    workbook: exeljs.Workbook,
    placeholders: { [key: string]: number },
    placeholderPlacement: Array<{ row: number, column: number, placeholder: string }>
) {
    const worksheet = workbook.getWorksheet(1);
    if (!worksheet) throw new Error("Worksheet not found in Excel file");

    for (const { row, column, placeholder } of placeholderPlacement) {
        const value = placeholders[placeholder];
        if (value !== undefined) {
            worksheet.getCell(row, column).value = value;
        }
    }
}

async function getTemplatePlaceholderPlacement(version: string) {
    let templateFiles = await fs.readdir("./templates");
    for (const file of templateFiles) {
        const workbook = new exeljs.Workbook();
        await workbook.xlsx.readFile(`./templates/${file}`);
        const fileVersion = await getFileVersion(workbook);
        if (fileVersion === version) {
            return await getPlaceholderPlacement(workbook);
        }
    }
    throw new Error(`Template file for version ${version} not found`);
}

async function getFileVersion(workbook: exeljs.Workbook) {
    const worksheet = workbook.getWorksheet(1);
    if (!worksheet) throw new Error("Worksheet not found in Excel file");
    let lastRow = -1;
    worksheet.eachRow((row, rowNumber) => lastRow = rowNumber);
    let latestVersion = worksheet.getCell(lastRow, 2).value?.toString().trim() || "";
    if (!latestVersion.startsWith("V")) throw new Error("Latest version not found in Excel file");
    return latestVersion.substring(1).trim(); // Remove the "V" prefix
}


async function getPlaceholderPlacement(workbook: exeljs.Workbook) {
    let placeholderPlacement: Array<{ row: number, column: number, placeholder: string }> = [];
    const worksheet = workbook.getWorksheet(1);
    if (!worksheet) throw new Error("Worksheet not found in template file");
    worksheet.eachRow((rowElement, row) => {
        rowElement.eachCell((cell, column) => {
            const placeholder = cell.value?.toString();
            if (!(placeholder?.startsWith("[") && placeholder.endsWith("]"))) return;
            placeholderPlacement.push({ row, column, placeholder });
        });
    });
    return placeholderPlacement;
}
