import { createRouter, createWebHistory } from 'vue-router';
import Home from '../views/Home.vue';
import Scripts from '../views/Scripts.vue';
import Settings from '../views/Settings.vue';

const routes = [
  {
    path: '/',
    component: Home,
    meta: {
      topBarTitle: '主页' // 设置顶部栏标题为 '主页'
    }
  },
  {
    path: '/scripts',
    component: Scripts,
    meta: {
      topBarTitle: '脚本' // 设置顶部栏标题为 '脚本'
    }
  },
  {
    path: '/settings',
    component: Settings,
    meta: {
      topBarTitle: '设置' // 设置顶部栏标题为 '设置'
    }
  }
];

const router = createRouter({
  history: createWebHistory(),
  routes
});

export default router;
