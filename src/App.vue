<template>
  <div id="app" class="flex-container">
    <aside class="sidebar">
      <div class="title">RPA Box</div>
      <router-link to="/" class="nav-link">主页</router-link>
      <router-link to="/scripts" class="nav-link">脚本</router-link>
      <router-link to="/settings" class="nav-link">设置</router-link>
    </aside>

    <main class="main-content">
      <router-view></router-view>
    </main>
  </div>
</template>

<script setup>
</script>

<style scoped>
/* 样式部分保持不变 */
</style>
