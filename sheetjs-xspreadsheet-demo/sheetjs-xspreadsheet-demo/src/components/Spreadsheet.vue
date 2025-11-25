<template>
  <div class="spreadsheet-wrapper">
    <!-- 调试模式下的控制面板 -->
    <div v-if="debug" class="toolbar">
      <!-- 基础配置 -->
      <div class="config-group">
        <div class="form-row">
             <label>行数: <input type="number" v-model.number="config.row.len" /></label>
        <label>列数: <input type="number" v-model.number="config.col.len" /></label>
        <label>
          Mode:
          <select v-model="config.mode">
            <option value="edit">edit</option>
            <option value="read">read</option>
          </select>
        </label>
        <label><input type="checkbox" v-model="config.showToolbar" /> 工具栏</label>
        <label><input type="checkbox" v-model="config.showGrid" /> 网格</label>
        <label><input type="checkbox" v-model="config.showContextmenu" /> 右键菜单</label>
        <label><input type="checkbox" v-model="config.showBottomBar" /> 底部栏</label>
        </div>
       

 <div class="form-row">  <label>行 <input type="number" v-model.number="cellRow" /></label>
        <label>列 <input type="number" v-model.number="cellCol" /></label>
        <label>值 <input type="text" v-model="cellText" /></label>
          <button @click="resetSpreadsheet">重置组件</button>
        <button @click="disableCell(cellRow, cellCol)">禁用单元格</button>
        <button @click="setCell(cellRow, cellCol, cellText)">设置单元格</button>
        <button @click="getCell(cellRow, cellCol)">获取单元格</button>
         <button @click="exportExcel">导出Excel</button>
        <input type="file" @change="importLocalFile" />
      </div>
        
      </div>

    </div>

    <!-- 表格容器 -->
    <div ref="spreadsheetContainer" class="spreadsheet"></div>
  </div>
</template>

<script setup>
import { ref, reactive, onMounted, watch } from "vue";
import Spreadsheet from "x-data-spreadsheet";
import ExcelJS from "exceljs";
import { stox, xtos } from "./excelspread";

const props = defineProps({
  mode: { type: String, default: "read" },
  showToolbar: { type: Boolean, default: false },
  showGrid: { type: Boolean, default: true },
  showContextmenu: { type: Boolean, default: false },
  showBottomBar: { type: Boolean, default: false },
  row: { type: Object, default: () => ({ len: 100, height: 25 }) },
  col: { type: Object, default: () => ({ len: 50, width: 100, indexWidth: 60, minWidth: 60 }) },
  debug: { type: Boolean, default: false },
});

const spreadsheetContainer = ref(null);
let spreadsheet = null;

// 配置响应式
const config = reactive({
  mode: props.mode,
  showToolbar: props.showToolbar,
  showGrid: props.showGrid,
  showContextmenu: props.showContextmenu,
  showBottomBar: props.showBottomBar,
  row: { ...props.row },
  col: { ...props.col },
});


// 公共参数
// 行列数浮动数量
const rowColFloat = ref(20);
const cellRow = ref(0);
const cellCol = ref(0);
const cellText = ref("");

// 初始化表格
function initSpreadsheet() {
  if (spreadsheetContainer.value) {
    spreadsheetContainer.value.innerHTML = "";
  }
  spreadsheet = new Spreadsheet(spreadsheetContainer.value, {
    mode: config.mode,
    showToolbar: config.showToolbar,
    showGrid: config.showGrid,
    showContextmenu: config.showContextmenu,
    showBottomBar: config.showBottomBar,
    row: config.row,
    col: config.col,
  });
}
function getMaxRowCol(sdata) {
  let maxRow = 0;
  let maxCol = 0;
  const sheet = sdata[0];

  Object.entries(sheet.rows || {}).forEach(([rIdx, row]) => {
    maxRow = Math.max(maxRow, parseInt(rIdx));
    Object.keys(row.cells || {}).forEach(cIdx => {
      maxCol = Math.max(maxCol, parseInt(cIdx));
    });
  });

  return { maxRow: maxRow + rowColFloat.value, maxCol: maxCol + rowColFloat.value }; // 加浮动2
}
onMounted(() => {
  initSpreadsheet();
});
function resetSpreadsheet() {
  // 恢复默认配置
  config.row.len = 100;
  config.col.len = 50;
  config.mode = "edit";
  config.showToolbar = true;
  config.showGrid = true;
  config.showContextmenu = true;
  config.showBottomBar = true;

  // 清空表格并重新初始化
  spreadsheetContainer.value.innerHTML = "";
  spreadsheet = new Spreadsheet(spreadsheetContainer.value, { ...config });
  spreadsheet.loadData([{ name: "Sheet1", rows: {}, cols: {} }]); // 空表
}
// ✅ 实时监听配置变化
watch(config, (newConfig) => {
  if (!spreadsheet) return;
  const data = spreadsheet.getData();
  spreadsheetContainer.value.innerHTML = "";
  spreadsheet = new Spreadsheet(spreadsheetContainer.value, newConfig);
  spreadsheet.loadData(data);
}, { deep: true });

// 导出 Excel
async function exportExcel() {
  const sdata = spreadsheet.getData();
  const wb = await xtos(sdata);
  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: "application/octet-stream" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "spreadsheet.xlsx";
  a.click();
  URL.revokeObjectURL(url);
}

// 本地上传文件导入
async function importLocalFile(event) {
  const file = event.target.files[0];
  if (!file) return;
  const wb = new ExcelJS.Workbook();
  const buffer = await file.arrayBuffer();
  await wb.xlsx.load(buffer);
  const sdata = await stox(wb);

  // 动态调整行列数：文件行列数 
  
  const { maxRow, maxCol } = getMaxRowCol(sdata);
config.row.len = maxRow;
config.col.len = maxCol;
console.log("import 数据解析:")
console.log(sdata)

  spreadsheet.loadData(sdata);
}


// 根据模板编码请求后端并渲染
async function loadTemplate(sdata) {
  if (!sdata) return;
  try {
   
    // 动态调整行列数：根据数据范围 + 浮动
    const { maxRow, maxCol } = getMaxRowCol(sdata);
    config.row.len = maxRow;
    config.col.len = maxCol;

    spreadsheet.loadData(sdata);
  } catch (err) {
    console.error("加载模板失败:", err);
  }
}

// 设置单元格
function setCell(row, col, text) {
  const data = spreadsheet.getData();
  if (!data[0].rows[row]) data[0].rows[row] = { cells: {} };
  data[0].rows[row].cells[col] = { text };
  spreadsheet.loadData(data);
}

// 获取单元格
function getCell(row, col) {
  const data = spreadsheet.getData();
  const cell = data[0].rows[row]?.cells[col];
  const valueObj={ row, col, text: cell ? cell.text : "" }
  alert(JSON.stringify(valueObj))
  console.log("获取单元格:", valueObj);
}

// 禁用单元格
function disableCell(row, col) {
  const data = spreadsheet.getData();
  if (!data[0].rows[row]) data[0].rows[row] = { cells: {} };
  if (!data[0].rows[row].cells[col]) data[0].rows[row].cells[col] = { text: "" };
  data[0].rows[row].cells[col].editable = false;
  spreadsheet.loadData(data);
}

defineExpose({
  loadTemplate
});

</script>

<style>
.spreadsheet-wrapper {
  width: 100%;
  height: 100%;
}
.spreadsheet {
  width: 100%;
  height: 600px;
}
.toolbar {
  margin-bottom: 10px;
}

input[type="number"], input[type="text"], select {
  width: 50px; /* ✅ 输入框宽度缩小到50 */
}


.form-row {
  display: flex;
  align-items: center;
  gap: 20px; /* 控制各元素之间的间距 */
   margin: 10px 0;  /* 上下外边距 10px */
}

label {
  font-weight: bold;
}


</style>
