import Vue from 'vue'
import Router from 'vue-router'
import Hello from '@/components/Hello'
import Index from '@/components/Index'
import A from '@/components/mian/a'
import B from '@/components/mian/b'
import C from '@/components/mian/c'
import D from '@/components/mian/d'
import ElementUI from 'element-ui'
// import 'element-ui/lib/theme-default/index.css'
import 'element-ui/lib/theme-chalk/index.css'
Vue.use(Router)
Vue.use(ElementUI)

import VueResource from 'vue-resource'
Vue.use(VueResource);
export default new Router({
  routes: [
    {
      path: '/',
      name: 'index',
      component: Index,
       children:[{
		      path: 'a',
		      name: 'A',
		      component: A
	    },{
		      path: 'b',
		      name: 'B',
		      component: B
	    },{
		      path: 'c',
		      name: 'C',
		      component: C
	    },{
		      path: 'd',
		      name: 'D',
		      component: D
	    }]
    }
    
  ]
})
