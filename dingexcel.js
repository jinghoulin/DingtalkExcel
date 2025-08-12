// sheet命名规则：系统-数据源/结果-sheet
const zentSourceSheet = Workbook.getSheet('禅道Bug数据源')
const jiraResultSheet = Workbook.getSheet('Jira重点Bug数据源')
const zentConditionSheet = Workbook.getSheet("禅道重点问题-筛选条件")
const allResultSheet = Workbook.getSheet("重点问题")
// bug原始数据的列数
const zentSourceColCount = 15;
// 禅道
const zentSourceFirstRow = zentSourceSheet.getRange('1:1')
const zentLevelColIndex = zentSourceFirstRow.find("严重程度").getColumn()
const zentTitleColIndex = zentSourceFirstRow.find("Bug标题").getColumn()
const zentKeywordColIndex = zentSourceFirstRow.find("关键词").getColumn()
const zentCreatdateColIndex = zentSourceFirstRow.find("创建日期").getColumn()
const zentSourceRowCount = zentSourceSheet.getRowCount();

const newDataArray = [];
const nowDate = new Date();
const colorRed = "#fe0300";
const colorOrange = "#fcc102";
const colorYellow = "#feff00";

// 清除筛选，以便读取到全部数据
if (allResultSheet.getFilter()) {
    allResultSheet.getFilter().delete();
}
if (zentSourceSheet.getFilter()) {
    zentSourceSheet.getFilter().delete();
}

// 给sheet添加筛选之前，先激活sheet。否则会报错"Cannot create a filter with other sheet's range"
allResultSheet.activate();
// 在第1行重新添加筛选，方便脚本运行后自行筛选
allResultSheet.filter('1:1');

// 添加筛选结果的标题行
zentSourceSheet.getRange(0, zentSourceColCount + 1, 1, 4).setValues([
    ["条件-严重程度", "条件-标题", "条件-关键词", "条件结果-或"],
]);

// 轮询zentSourceSheet的每一行
const zentNotNullRowCount = getNotNullRowCount(zentSourceSheet);
var isMatchKeyRules = false;
for (let i = 1; i < zentNotNullRowCount; i++) {
    // 匹配重点问题，并获取匹配结果
    isMatchKeyRules = matchKeyRules(zentSourceSheet, i);

    // 若allResultSheet中存在此BugId，则直接更新数据
    const bugId = zentSourceSheet.getRange(i, 0, 1, 1).getValue()
    const findRange = allResultSheet.getRange(0, 0, allResultSheet.getRowCount(), 1).find(bugId.toString())
    if (findRange) {
        if (!isMatchKeyRules) {// 删除allResultSheet中不符合重点问题条件的行
            if (findRange.getRow() != 0) {// 第0行是标题行
                allResultSheet.deleteRow(findRange.getRow());
                Output.log(`allResultSheet delete row, bugId=${bugId}`)
            }
        } else {// 更新allResultSheet中符合重点问题条件的行
            updateRow(zentSourceSheet, i, allResultSheet, findRange.getRow());
            Output.log(`allResultSheet update zentao row, bugId=${bugId}`)
        }
    } else {
        if (isMatchKeyRules) {
            // 将allResultSheet中不存在的新数据保存到数组
            addRowToArray(zentSourceSheet, i);
            // Output.log(`added zentao new data, array = ${newDataArray}`)
        }
    }
}

// 轮询keyBugJiraSheet每一行
const jiraCount = getNotNullRowCount(jiraResultSheet);
for (let i = 1; i < jiraCount; i++) {
    // 若allResultSheet中存在此BugId，则直接更新数据
    const bugId = jiraResultSheet.getRange(i, 0, 1, 1).getValue()
    const findRange = allResultSheet.getRange(0, 0, allResultSheet.getRowCount(), 1).find(bugId.toString())
    if (findRange) {
        // 更新allResultSheet中符合重点问题条件的行
        updateRow(jiraResultSheet, i, allResultSheet, findRange.getRow(), true);
        Output.log(`allResultSheet update jira row, bugId=${bugId}`)
    } else {
        // 将allResultSheet中不存在的新数据保存到数组
        addRowToArray(jiraResultSheet, i, true);
        // Output.log(`added jira new data, array = ${newDataArray}`)
    }
}

// 将暂存的所有行追加到allResultSheet
appendArrayToSheet(allResultSheet);

// 轮询allResultSheet的每一行
// for (let i = 1; i < getNotNullRowCount(allResultSheet); i++) {
//     // 添加存活时间颜色
//     addAliveColors(allResultSheet, i);
// }

zentSourceSheet.setRowsHeight(0, zentSourceSheet.getRowCount(), 22);
allResultSheet.setRowsHeight(0, allResultSheet.getRowCount(), 22);


// -------------------------functions-------------------------

function appendArrayToSheet(targetSheet) {
    if (newDataArray.length === 0) {
        Output.log("没有新增的（重点问题）数据");
    } else {
        Output.log("filteredData lenthg=" + newDataArray.length)
        appendValues(newDataArray, targetSheet)
    }
}

function addRowToArray(sourceSheet, sourceRowIndex, isJira) {
    // 获取一行的数据
    const rowData = sourceSheet.getRange(sourceRowIndex, 0, 1, zentSourceColCount).getValues()[0];

    if (isJira) {
        // 调整column与allResultSheet结构一致
        rowData[10] = rowData[6];
        rowData[7] = rowData[5];
        rowData[6] = rowData[4];
        rowData[5] = rowData[3];
        rowData[3] = rowData[2];
        // 清空无效column
        rowData[4] = '';
        rowData[2] = '';
    }

    newDataArray.push(rowData);
}

// 添加重点问题规则的判定结果，并return
function matchKeyRules(sheet, index) {
    var result = false;

    // 严重程度
    const levelValue = sheet.getRange(index, zentLevelColIndex, 1, 1).getValue();
    if (levelValue) {
        const levelResult = sheet.getRange(index, zentLevelColIndex, 1, 1).getValue() <= zentConditionSheet.getRange('A2').getValue()
        sheet.getRange(index, zentSourceColCount + 1, 1, 1).setValue(levelResult)
        if (levelResult) result = true;
    } else {
        sheet.getRange(index, zentSourceColCount + 1, 1, 1).setValue(false)
    }
    // 标题
    const titleValue = sheet.getRange(index, zentTitleColIndex, 1, 1).getValue()
    if (titleValue) {
        var titleResult = false;
        const titleCondition = zentConditionSheet.getRange('B2').getValue().toString().split("，");
        if (titleCondition.find(item => titleValue.includes(item))) {
            titleResult = true;
        }
        sheet.getRange(index, zentSourceColCount + 2, 1, 1).setValue(titleResult)
        if (titleResult) result = true;
    } else {
        sheet.getRange(index, zentSourceColCount + 2, 1, 1).setValue(false)
    }
    // 关键词
    const keywordValue = sheet.getRange(index, zentKeywordColIndex, 1, 1).getValue()
    if (keywordValue) {
        var keywordResult = false;
        const keywordCondition = zentConditionSheet.getRange('C2').getValue().toString().split("，");
        if (keywordCondition.find(item => keywordValue.includes(item))) {
            keywordResult = true;
        }
        sheet.getRange(index, zentSourceColCount + 3, 1, 1).setValue(keywordResult)
        if (keywordResult) result = true;
    } else {
        sheet.getRange(index, zentSourceColCount + 3, 1, 1).setValue(false)
    }
    // 记录最终判定结果
    sheet.getRange(index, zentSourceColCount + 4, 1, 1).setValue(result)

    return result;
}

// 根据bug存活时间添加颜色
function addAliveColors(sheet, index) {
    // 计算存活时间。空白行的日期为1970-01-01
    const createDateValue = new Date(sheet.getRange(index, zentCreatdateColIndex, 1, 1).getValue());
    const aliveDays = parseInt((nowDate - createDateValue) / (1000 * 60 * 60 * 24));
    // Output.log(`index=${index}, createDateValue=${createDateValue}, aliveDays=${aliveDays}`)
    if (aliveDays > 15) {
        sheet.getRange(index, zentCreatdateColIndex, 1, 1).setBackgroundColor(colorRed)
    } else if (aliveDays > 10) {
        sheet.getRange(index, zentCreatdateColIndex, 1, 1).setBackgroundColor(colorOrange)
    } else if (aliveDays > 5) {
        sheet.getRange(index, zentCreatdateColIndex, 1, 1).setBackgroundColor(colorYellow)
    }
}

// 追加数据到空行
function appendValues(values, targetSheet) {
    targetSheet.getRange(getNotNullRowCount(targetSheet), 0, values.length, values[0].length).setValues(values, { parseType: 'raw' })
    Output.log(`共复制 ${values.length} 行数据到 "${allResultSheet.getName()}"`);
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
function updateRow(sourceSheet, sourceRowIndex, targetSheet, targetRowIndex, isJira) {
    const targetRange = targetSheet.getRange(targetRowIndex, 0, 1, zentSourceColCount)
    const sourceRange = sourceSheet.getRange(sourceRowIndex, 0, 1, zentSourceColCount)
    if (isJira) {
        // Output.log(`updateRow: sourceRange.getValues() = ${sourceRange.getValues()}`)
        const rowData = sourceRange.getValues()[0];

        const newArray = new Array(1);
        newArray[0] = new Array(zentSourceColCount);
        //调整column与allResultSheet结构一致
        newArray[0][14] = "Jira";
        newArray[0][10] = rowData[6];
        newArray[0][7] = rowData[5];
        newArray[0][6] = rowData[4];
        newArray[0][5] = rowData[3];
        newArray[0][3] = rowData[2];
        newArray[0][1] = rowData[1];
        newArray[0][0] = rowData[0];

        targetRange.setValues(newArray, { parseType: 'raw' })
    } else {
        targetRange.setValues(sourceRange.getValues(), { parseType: 'raw' })
    }

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
// // allResultSheet.getRange(rangeValue).setValues(bugSheet.getActiveRange().getValues())

// const startRow = 0;

// // 遍历所有行，筛选出未被隐藏的行（即筛选结果中的行）
// for (let row = startRow; row < rowCount; row++) {
//     if (bugSheet.getRowVisibility(row)=='visible') {
//         const bugIdValue = bugSheet.getRange(row, 0, 1, colCount).getValue()
//         // 跳过Bug编号不是数字的行
//         if(isNaN(bugIdValue)){
//             continue;
//         }
//         const findRange = allResultSheet.getRange(0, 0, rowCount, 1).find(bugIdValue.toString())
//         // allResultSheet中若第1列中存在此BugId，则直接更新到此行
//         if(findRange){
//             updateRow(bugSheet, row, allResultSheet, findRange.getRow());
//             continue;
//         }
//         // 保存allResultSheet中不存在的新数据
//         const rowData = bugSheet.getRange(row, 0, 1, colCount).getValues()[0];
//         // Output.log(rowData)
//         filteredData.push(rowData);
//     }
// }