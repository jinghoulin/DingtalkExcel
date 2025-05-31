// const originSheet = Workbook.getSheet("bug原始数据");

// const originRange = originSheet.filter('A:N');

// const filterCondition1 = {
//     operator: 'less-equal',
//     value: 'B',
// }
// const criteria = Workbook.newFilterCriteriaBuilder().setVisibleConditions(filterCondition1).build();

// const filter = originSheet.getFilter();
// filter.setColumnFilterCriteria(3, criteria);

const bugSheet = Workbook.getSheet('禅道bug数据')
const conditionSheet = Workbook.getSheet("重点问题-筛选条件")
const keyBugSheet = Workbook.getSheet("重点问题")

// bugSheet.setRowsHeight(0,600,20)
// bugSheet.setFrozenRowCount(1)

const firstRow = bugSheet.getRange('1:1')

var filter=bugSheet.getFilter()
if(filter){
    bugSheet.deleteFilter()
}
filter = bugSheet.filter(firstRow)

const levelColumnIndex = firstRow.find("严重程度").getColumn()
Output.log(levelColumnIndex)

const value1 = conditionSheet.getRange('A2').getValue().toString()
Output.log(value1)

const filterCondition1 = {operator: 'less-equal', value: value1,}
const filterCondition2 = {}
// const criteria = Workbook.newFilterCriteriaBuilder().setVisibleConditions(filterCondition1, filterCondition2, 'or').build()
const criteria = Workbook.newFilterCriteriaBuilder().setVisibleConditions(filterCondition1).build()
filter.setColumnFilterCriteria(levelColumnIndex, criteria)


// 筛选后复制
const rangeValue = 'A1:N600'
// bugSheet.setActiveRange(rangeValue)
// Output.log(filter.getRange())
// Output.log(bugSheet.getActiveRangeList())
// keyBugSheet.getRange(rangeValue).setValues(bugSheet.getActiveRange().getValues())


// 假设数据从第1行开始，A1是标题行
const startRow = 1;
const endRow = 600;       // 获取最大行号
const colCount = 14;  // 获取列数

const filteredData = [];

// 遍历所有行，筛选出未被隐藏的行（即筛选结果中的行）
for (let row = startRow; row < endRow; row++) {
    if (bugSheet.getRowVisibility(row)=='visible') {
        const bugIdValue = bugSheet.getRange(row, 0, 1, colCount).getValue()
        // 跳过Bug编号不是数字的行
        if(isNaN(bugIdValue)){
            continue;
        }
        const findRange = keyBugSheet.getRange(0, 0, endRow, 1).find(bugIdValue.toString())
        // keyBugSheet中若第1列中存在此BugId，则直接更新到此行
        if(findRange){
            updateRow(bugSheet, row, keyBugSheet, findRange.getRow());
            continue;
        }
        // 保存keyBugSheet中不存在的新数据
        const rowData = bugSheet.getRange(row, 0, 1, colCount).getValues()[0];
        // Output.log(rowData)
        filteredData.push(rowData);
    }
}


if (filteredData.length === 0) {
    Output.log("没有符合条件的数据");
    // return;
}else{
    Output.log("filteredData lenthg="+filteredData.length)
    appendValues(filteredData, keyBugSheet)
    
}

// 追加数据到空行
function appendValues(values, targetSheet){
    targetSheet.getRange(firstNullRowIndex(targetSheet), 0, filteredData.length, filteredData[0].length).setValues(filteredData)
    Output.log(`共复制 ${filteredData.length} 行数据到 "${keyBugSheet.getName()}"`);
}


// 将筛选后的数据写入目标 Sheet
// keyBugSheet.getRange(1, 0, filteredData.length, filteredData[0].length).setValues(filteredData);




// 获取第一个空行
function firstNullRowIndex(sheet){
    var index = 0;
    for(; index <endRow; index++){
        // 第1列和第2列都为null，则判定为空行
        if(sheet.getRange(index, 0, 1, 1).getValue()==null && sheet.getRange(index, 1, 1, 1).getValue()==null){
            break;
        }
    }
    Output.log("firstNullRowIndex: "+index)
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

