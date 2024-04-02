<template>
  <div class="top-bar">
    <div class="title">{{ topBarTitle }}</div>
    <button type="button" class="btn" @click="addScript">添加脚本</button>
  </div>
  <div class="content">
    <!-- <h2>脚本管理器</h2> -->
    <div>
      
      <ul>
        <li  class="card" v-for="script in scripts" :key="script.id">
          <span class="scriptname">{{ script.name }}</span><button class="btn" @click="runScripts(script.file)">运行脚本</button>
        </li>
        
      </ul>
      
    </div>
  </div>

</template>

<script setup lang="ts">
import { defineComponent } from 'vue';
import { invoke } from '@tauri-apps/api/tauri';
let topBarTitle = "脚本管理";

let scripts = [{
    id: 1,
    name: "处理发货清单",
    file: "../script/处理发货清单/main.py",
  },
  {
    id: 2,
    name: "营业收入分析",
    file: "../script/营业收入分析/excel_to_word.py",
  }

];


function addScript() {
  // 添加脚本的逻辑
}

function runScripts(file) {
  // 运行脚本的逻辑
  invoke('run_python_script',{ file: file })
  .then((response) => console.log(response))
  
}
</script>

<style scoped>
/* 你的组件样式 */
</style>