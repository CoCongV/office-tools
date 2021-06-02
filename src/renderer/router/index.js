import Vue from 'vue'
import Router from 'vue-router'

Vue.use(Router)

export default new Router({
  routes: [
    {
      path: '/',
      name: 'home',
      component: require('@/views/Home').default
    },
    {
      path: '/excelMatch',
      name: 'excelMatch',
      component: () => import('@/views/ExcelMatch.vue')
    },
    {
      path: '/excelFill',
      name: 'excelFill',
      component: () => import('@/views/ExcelFill.vue')
    }
  ]
})
