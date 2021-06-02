<template>
  <div id="app">
    <a-layout id="components-layout-demo-side" style="min-height: 100vh">
    <a-layout-sider v-model="collapsed" collapsible>
      <div class="logo" />
      <a-menu theme="dark" :default-selected-keys="['0']" mode="inline">
        <a-menu-item key="0" @click="$router.push({name: 'home'})">
          <a-icon type="home" />
          <span>主页</span>
        </a-menu-item>
        <a-sub-menu key="sub1">
          <span slot="title"><a-icon type="file-excel" /><span>Excel</span></span>
          <a-menu-item key="1" @click="$router.push({name: 'excelMatch'})">
            多表查询匹配数据
          </a-menu-item>
          <a-menu-item key="2" @click="$router.push({name: 'excelFill'})">
            多表查询填充数据
          </a-menu-item>
        </a-sub-menu>
      </a-menu>
    </a-layout-sider>
    <a-layout>
      <a-layout-header style="background: #fff; padding: 0" />
      <a-layout-content style="margin: 0 16px">
        <router-view></router-view>
      </a-layout-content>
      <a-layout-footer style="text-align: center">
        Office Tools ©2020 Created by CoCong
      </a-layout-footer>
    </a-layout>
  </a-layout>
  </div>
</template>

<script>
export default {
  name: "office-tools",
  data() {
    return {
      collapsed: false,
    };
  },
};
</script>

<style>
/* CSS */
</style>
