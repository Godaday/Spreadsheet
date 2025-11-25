<!--
 * @description: -
 * @version: 1.0
 * @Author: yanwf1
 * @Date: 2025-11-22 11:42:14
 * @LastEditors: -
 * @LastEditTime: 2025-11-23 19:51:08
-->
<template>
  <div id="app">
    <h2>Spreadsheet 组件测试</h2>
    <!-- 引用封装好的 Spreadsheet 组件 -->
   
<div class="form-row">
  <label for="templateCode">模板 Code</label>
  <input id="templateCode" v-model="excelTemplateCode" />
  <button @click="loadServerTemplate" :disabled="loading">{{ loading?'请求模板数据....':'加载模板' }}</button>

  <label for="debug">调试模式</label>
  <input id="debug" type="checkbox" v-model="debug" />
</div>
<Spreadsheet
  ref="sheet"
      :mode="mode"
      :showToolbar="showToolbar"
      :showGrid="showGrid"
      :showContextmenu="showContextmenu"
      :showBottomBar="showBottomBar"
      :debug="debug"
    />
  </div>


</template>

<script setup>
import { ref } from "vue";
import Spreadsheet from "./components/Spreadsheet.vue";
const sheet = ref(null);
const excelTemplateCode=ref('')
const mode = ref("edit");
const debug = ref(false);
const showToolbar = ref(true);
const showGrid = ref(true);
const showContextmenu = ref(true);
const showBottomBar = ref(true);
const loading= ref(false)

const loadServerTemplate=async ()=> {
  // 假设模板编码是 "TMP001"
  if(!excelTemplateCode.value)
{
   alert('excelTemplateCode is null')
   
}
else{



  try {
     loading.value = true;
    const resp = await fetch(`https://localhost:9001/api/template/${excelTemplateCode.value}`);
    if (!resp.ok) throw new Error("请求失败");
    const sdata = await resp.json();
 sheet.value.loadTemplate(sdata.data);

  } catch (err) {
    console.error("加载模板失败:", err);
    alert("加载模板失败:"+ err)
  }
    finally {
       loading.value = false;
      }


}
 
}

</script>

<style scoped>
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
