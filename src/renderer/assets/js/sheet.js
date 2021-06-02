import XLSX from "xlsx";
const { dialog } = require("electron").remote;

export function dealData(titles, rows) {
    let matchFile = titles[0];
    let matchKeys = [];
    let matchedKeys = [];
    for (let i = 0; i < rows; i++) {
        matchKeys.push(matchFile.get("match" + i));
        matchedKeys.push(titles[1].get("match" + i));
    }
    // 处理matchFile 数据
    let cacheMatchValue = [];
    for (let valueMatchFile of matchFile.get("data")) {
        let cache = [];
        let cacheStr = "";
        for (let matchKey of matchKeys) {
            cache.push(valueMatchFile[matchKey].trim());
            // TODO: upper
            // cache[matchIndex] =  valueMatchFile[matchKeys[matchIndex]]
        }
        cacheStr = JSON.stringify(cache);
        cacheMatchValue.push(cacheStr);
    }
    // 处理matchedFile 数据
    let cacheMatchedSheetList = [];
    for (let title of titles.slice(1)) {
        let cacheMatchedSheet = []
        for (let valueMatchedFile of title.get("data")) {
            let cache = [];
            let cacheStr = "";
            for (let matchedKey of matchedKeys) {
                cache.push(valueMatchedFile[matchedKey]);
            }
            cacheStr = JSON.stringify(cache);
            cacheMatchedSheet.push(cacheStr);
        }
        cacheMatchedSheetList.push(cacheMatchedSheet)
    }
    return [cacheMatchValue, cacheMatchedSheetList];
};

export async function save(data) {
    let ws = XLSX.utils.json_to_sheet(data)
    let wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "sheet1")
    let filePath = await dialog.showSaveDialog({
        defaultPath: "generate.xlsx",
    })
    if (filePath) {
        XLSX.writeFile(wb, filePath);
    }
}
