const bugSheet = Workbook.getSheet('禅道bug数据')
const conditionSheet = Workbook.getSheet("重点问题-筛选条件")
const keyBugSheet = Workbook.getSheet("重点问题")

const keyRulesColumnIndex = 24;// 禅道bug的默认模版有23列

const bugFirstRow = bugSheet.getRange('1:1')
const levelColumnIndex = bugFirstRow.find("严重程度").getColumn()
const titleColumnIndex = bugFirstRow.find("Bug标题").getColumn()
const keywordColumnIndex = bugFirstRow.find("关键词").getColumn()

const filteredData = [];
const rowCount = bugSheet.getRowCount();
const colCount = 14;

addKeyRules()

// 每行记录符合重点问题的判定结果
function addKeyRules(){
    const bugRowCount = firstNullRowIndex(bugSheet);
    for(let i=1; i<bugRowCount; i++){
        var result = false;

        // 严重程度
        const levelValue = bugSheet.getRange(i, levelColumnIndex, 1, 1).getValue();
        if(levelValue){
            const levelResult = bugSheet.getRange(i, levelColumnIndex, 1, 1).getValue() <= conditionSheet.getRange('A2').getValue()
            bugSheet.getRange(i, keyRulesColumnIndex, 1, 1).setValue(levelResult)
            if(levelResult) result = true;
        }else{
            bugSheet.getRange(i, keyRulesColumnIndex, 1, 1).setValue(false)
        }
        // 标题
        const titleValue = bugSheet.getRange(i, titleColumnIndex, 1, 1).getValue()
        if(titleValue){
            var titleResult = false;
            const titleCondition = conditionSheet.getRange('B2').getValue().toString().split("，");
            if(titleCondition.find(item => titleValue.includes(item))){
                titleResult = true;
            }
            bugSheet.getRange(i, keyRulesColumnIndex+1, 1, 1).setValue(titleResult)
            if(titleResult) result = true;
        }else{
            bugSheet.getRange(i, keyRulesColumnIndex+1, 1, 1).setValue(false)
        }
        // 关键词
        const keywordValue = bugSheet.getRange(i, keywordColumnIndex, 1, 1).getValue()
        if(keywordValue){
            var keywordResult = false;
            const keywordCondition = conditionSheet.getRange('C2').getValue().toString().split("，");
            if(keywordCondition.find(item => keywordValue.includes(item))){
                keywordResult = true;
            }
            bugSheet.getRange(i, keyRulesColumnIndex+2, 1, 1).setValue(keywordResult)
            if(keywordResult) result = true;
        }else{
            bugSheet.getRange(i, keyRulesColumnIndex+2, 1, 1).setValue(false)
        }

        // 记录最终判定结果
        bugSheet.getRange(i, keyRulesColumnIndex+3, 1, 1).setValue(result)
        if(!result) continue;
        
        const bugIdValue = bugSheet.getRange(i, 0, 1, 1).getValue()
        const findRange = keyBugSheet.getRange(0, 0, rowCount, 1).find(bugIdValue.toString())
        // keyBugSheet中若第1列中存在此BugId，则直接更新到此行
        if(findRange){
            updateRow(bugSheet, i, keyBugSheet, findRange.getRow());
            continue;
        }
        // 保存keyBugSheet中不存在的新数据
        const rowData = bugSheet.getRange(i, 0, 1, colCount).getValues()[0];
        // Output.log(rowData)
        filteredData.push(rowData);
    }

}

if (filteredData.length === 0) {
    Output.log("没有新增的（重点问题）数据");
    // return;
}else{
    Output.log("filteredData lenthg="+filteredData.length)
    appendValues(filteredData, keyBugSheet)
}

keyBugSheet.setRowsHeight(0, keyBugSheet.getRowCount(), 22);

// 追加数据到空行
function appendValues(values, targetSheet){
    targetSheet.getRange(firstNullRowIndex(targetSheet), 0, filteredData.length, filteredData[0].length).setValues(filteredData)
    Output.log(`共复制 ${filteredData.length} 行数据到 "${keyBugSheet.getName()}"`);
}

// 获取第一个空行
function firstNullRowIndex(sheet){
    var index = 0;
    for(; index < sheet.getRowCount(); index++){
        // 第1列和第2列都为null，则判定为空行
        if(sheet.getRange(index, 0, 1, 1).getValue()==null && sheet.getRange(index, 1, 1, 1).getValue()==null){
            break;
        }
    }
    Output.log(sheet.getName()+" firstNullRowIndex: "+index)
    return index;
}

// 将1行更新到指定行
function updateRow(sourceSheet, sourceRowIndex, targetSheet, targetRowIndex){
    const targetRange = targetSheet.getRange(targetRowIndex, 0, 1, colCount)
    const sourceRange = sourceSheet.getRange(sourceRowIndex, 0, 1, colCount)
    targetRange.setValues(sourceRange.getValues())
    // Output.log("updateRow: "+targetSheet.getName()+"->"+targetRowIndex)
}

// 打印对象的类型（钉钉API很多没有重写toString）
function printType(obj){
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

// bugSheet.setRowsHeight(0,600,20)
// bugSheet.setFrozenRowCount(1)

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