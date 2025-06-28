const bugSheet = Workbook.getSheet('禅道bug数据')
const conditionSheet = Workbook.getSheet("重点问题-筛选条件")
const keyBugSheet = Workbook.getSheet("重点问题")
// bug原始数据的列数
const bugOriginColCount = 14;

const bugFirstRow = bugSheet.getRange('1:1')
const levelColumnIndex = bugFirstRow.find("严重程度").getColumn()
const titleColumnIndex = bugFirstRow.find("Bug标题").getColumn()
const keywordColumnIndex = bugFirstRow.find("关键词").getColumn()
const creatDateColumnIndex = bugFirstRow.find("创建日期").getColumn()

const newDataArray = [];
const rowCount = bugSheet.getRowCount();
const nowDate = new Date();
const colorRed = "#fe0300";
const colorOrange = "#fcc102";
const colorYellow = "#feff00";

// 清除筛选，才能读取到全部数据
if(keyBugSheet.getFilter()){
    keyBugSheet.getFilter().delete();
}
if(bugSheet.getFilter()){
    bugSheet.getFilter().delete();
}
// 在第1行重新添加筛选，方便脚本运行后自行筛选
keyBugSheet.filter('1:1');

// 轮询sourceSheet的每一行
const bugRowCount = getNotNullRowCount(bugSheet);
var isMatchKeyRules = false;
for (let i = 1; i < bugRowCount; i++) {
    // 匹配重点问题，并获取匹配结果
    isMatchKeyRules = matchKeyRules(bugSheet, i);

    // 若keyBugSheet中存在此BugId，则直接更新数据
    const bugId = bugSheet.getRange(i, 0, 1, 1).getValue()
    const findRange = keyBugSheet.getRange(0, 0, rowCount, 1).find(bugId.toString())
    if (findRange) {
        if (!isMatchKeyRules) {// 删除keyBugSheet中不符合重点问题条件的行
            if (findRange.getRow() != 0) {// 第0行是标题行
                keyBugSheet.deleteRow(findRange.getRow());
                Output.log(`keyBugSheet delete row, bugId=${bugId}`)
            }
        } else {// 更新keyBugSheet中符合重点问题条件的行
            updateRow(bugSheet, i, keyBugSheet, findRange.getRow());
            Output.log(`keyBugSheet update row, bugId=${bugId}`)
        }
    } else {
        if (isMatchKeyRules) {
            // 将一行数据暂存到数组
            addRowToArray(bugSheet, i, keyBugSheet);
        }
    }
}
// 将暂存的所有行追加到keyBugSheet
appendArrayToSheet(keyBugSheet);

// 轮询targetSheet的每一行
var keyBugRowCount = getNotNullRowCount(keyBugSheet);
for (let i = 1; i < keyBugRowCount; i++) {
    // 添加存活时间颜色
    addAliveColors(keyBugSheet, i);
}

bugSheet.setRowsHeight(0, bugSheet.getRowCount(), 22);
keyBugSheet.setRowsHeight(0, keyBugSheet.getRowCount(), 22);
// -------------------------functions-------------------------

function appendArrayToSheet(targetSheet) {
    if (newDataArray.length === 0) {
        Output.log("没有新增的（重点问题）数据");
    } else {
        Output.log("filteredData lenthg=" + newDataArray.length)
        appendValues(newDataArray, targetSheet)
    }
}

function addRowToArray(sourceSheet, sourceRowIndex) {
    // 保存keyBugSheet中不存在的新数据
    const rowData = sourceSheet.getRange(sourceRowIndex, 0, 1, bugOriginColCount).getValues()[0];
    newDataArray.push(rowData);
    // Output.log(rowData)
}

// 添加重点问题规则的判定结果，并return
function matchKeyRules(sheet, index) {
    var result = false;

    // 严重程度
    const levelValue = sheet.getRange(index, levelColumnIndex, 1, 1).getValue();
    if (levelValue) {
        const levelResult = sheet.getRange(index, levelColumnIndex, 1, 1).getValue() <= conditionSheet.getRange('A2').getValue()
        sheet.getRange(index, bugOriginColCount, 1, 1).setValue(levelResult)
        if (levelResult) result = true;
    } else {
        sheet.getRange(index, bugOriginColCount, 1, 1).setValue(false)
    }
    // 标题
    const titleValue = sheet.getRange(index, titleColumnIndex, 1, 1).getValue()
    if (titleValue) {
        var titleResult = false;
        const titleCondition = conditionSheet.getRange('B2').getValue().toString().split("，");
        if (titleCondition.find(item => titleValue.includes(item))) {
            titleResult = true;
        }
        sheet.getRange(index, bugOriginColCount + 1, 1, 1).setValue(titleResult)
        if (titleResult) result = true;
    } else {
        sheet.getRange(index, bugOriginColCount + 1, 1, 1).setValue(false)
    }
    // 关键词
    const keywordValue = sheet.getRange(index, keywordColumnIndex, 1, 1).getValue()
    if (keywordValue) {
        var keywordResult = false;
        const keywordCondition = conditionSheet.getRange('C2').getValue().toString().split("，");
        if (keywordCondition.find(item => keywordValue.includes(item))) {
            keywordResult = true;
        }
        sheet.getRange(index, bugOriginColCount + 2, 1, 1).setValue(keywordResult)
        if (keywordResult) result = true;
    } else {
        sheet.getRange(index, bugOriginColCount + 2, 1, 1).setValue(false)
    }
    // 记录最终判定结果
    sheet.getRange(index, bugOriginColCount + 3, 1, 1).setValue(result)

    return result;
}

// 根据bug存活时间添加颜色
function addAliveColors(sheet, index) {
    // 计算存活时间。空白行的日期为1970-01-01
    const createDateValue = new Date(sheet.getRange(index, creatDateColumnIndex, 1, 1).getValue());
    const aliveDays = parseInt((nowDate - createDateValue) / (1000 * 60 * 60 * 24));
    // Output.log(`index=${index}, createDateValue=${createDateValue}, aliveDays=${aliveDays}`)
    if (aliveDays > 15) {
        sheet.getRange(index, creatDateColumnIndex, 1, 1).setBackgroundColor(colorRed)
    } else if (aliveDays > 10) {
        sheet.getRange(index, creatDateColumnIndex, 1, 1).setBackgroundColor(colorOrange)
    } else if (aliveDays > 5) {
        sheet.getRange(index, creatDateColumnIndex, 1, 1).setBackgroundColor(colorYellow)
    }
}

// 追加数据到空行
function appendValues(values, targetSheet) {
    targetSheet.getRange(getNotNullRowCount(targetSheet), 0, values.length, values[0].length).setValues(values, { parseType: 'raw' })
    Output.log(`共复制 ${values.length} 行数据到 "${keyBugSheet.getName()}"`);
}

// 获取有效行的count，也是第一个空行的index
function getNotNullRowCount(sheet) {
    var count = 0;
    for (; count < sheet.getRowCount(); count++) {
        // 第1列和第2列都为null，则判定为空行
        if (sheet.getRange(count, 0, 1, 1).getValue() == null && sheet.getRange(count, 1, 1, 1).getValue() == null) {
            break;
        }
    }
    Output.log(`getNotNullRowCount: ${sheet.getName()}sheet 有效行数=${count}`);
    return count;
}

// 将1行更新到指定行
function updateRow(sourceSheet, sourceRowIndex, targetSheet, targetRowIndex) {
    const targetRange = targetSheet.getRange(targetRowIndex, 0, 1, bugOriginColCount)
    const sourceRange = sourceSheet.getRange(sourceRowIndex, 0, 1, bugOriginColCount)
    targetRange.setValues(sourceRange.getValues(), { parseType: 'raw' })
    // Output.log("updateRow: "+targetSheet.getName()+"->"+targetRowIndex)
}

// 打印对象的类型（钉钉API很多没有重写toString）
function printType(obj) {
    Output.log(Object.prototype.toString.call(obj))
}



// const originSheet = Workbook.getSheet("bug原始数据");

// const originRange = originSheet.filter('A:N');

// const filterCondition1 = {
//     operator: 'less-equal',
//     value: 'B',
// }
// const criteria = Workbook.newFilterCriteriaBuilder().setVisibleConditions(filterCondition1).build();

// const filter = originSheet.getFilter();
// filter.setColumnFilterCriteria(3, criteria);

// var filter=bugSheet.getFilter()
// if(filter){
//     bugSheet.deleteFilter()
// }
// filter = bugSheet.filter(bugFirstRow)

// const value1 = conditionSheet.getRange('A2').getValue()

// const filterCondition1 = {operator: 'less-equal', value: value1,}
// const filterCondition2 = {}
// // const criteria = Workbook.newFilterCriteriaBuilder().setVisibleConditions(filterCondition1, filterCondition2, 'or').build()
// const criteria = Workbook.newFilterCriteriaBuilder().setVisibleConditions(filterCondition1).build()
// filter.setColumnFilterCriteria(levelColumnIndex, criteria)

// // 筛选后复制
// const rangeValue = 'A1:N600'
// // bugSheet.setActiveRange(rangeValue)
// // Output.log(filter.getRange())
// // Output.log(bugSheet.getActiveRangeList())
// // keyBugSheet.getRange(rangeValue).setValues(bugSheet.getActiveRange().getValues())

// const startRow = 0;

// // 遍历所有行，筛选出未被隐藏的行（即筛选结果中的行）
// for (let row = startRow; row < rowCount; row++) {
//     if (bugSheet.getRowVisibility(row)=='visible') {
//         const bugIdValue = bugSheet.getRange(row, 0, 1, colCount).getValue()
//         // 跳过Bug编号不是数字的行
//         if(isNaN(bugIdValue)){
//             continue;
//         }
//         const findRange = keyBugSheet.getRange(0, 0, rowCount, 1).find(bugIdValue.toString())
//         // keyBugSheet中若第1列中存在此BugId，则直接更新到此行
//         if(findRange){
//             updateRow(bugSheet, row, keyBugSheet, findRange.getRow());
//             continue;
//         }
//         // 保存keyBugSheet中不存在的新数据
//         const rowData = bugSheet.getRange(row, 0, 1, colCount).getValues()[0];
//         // Output.log(rowData)
//         filteredData.push(rowData);
//     }
// }