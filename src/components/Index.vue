<template>
  <div class="index">
	   <goodsd v-on:mess="mess"></goodsd>
	   <div id="nav">
	   	<ul>
	   		<li><router-link to="/a">首页</router-link></li>
	   		<li><router-link to="/b">产品页</router-link></li>
	   		<li><router-link to="/c">内容页</router-link></li>
	   		<li><router-link to="/d">内容页</router-link></li>
	   		<div>
	   			{{age}}
	   		</div>
	   	</ul>
	   </div>
   	 <div class="main">
   	 		<router-view></router-view>
   	 </div>
   	 <div class="footer">
   	 		 <foo message="zhngsan"></foo>
   	 </div>
   	 
  </div>
</template>

<script>
import goodsd from '@/components/goodsd';
import foo from '@/components/footer/foo';

import Vue from 'vue';
export default {
  name: 'app',
  components:{
   goodsd,
   foo
  },
  data() {
      return {   
      	age:"今年了26"
      }
  },
  methods:{

  	mess(mes){
  		console.log(mes)
  	}
  }
}
</script>

<!-- Add "scoped" attribute to limit CSS to this component only -->
<style lang="less">
.index{
	width: 100%;
	height: 100%;
	background: #ccc;
}
#nav ul{
		overflow: hidden;
	}
#nav ul li{
		float: left;
		font: 24px/30px "微软雅黑";
		color: #000;
		margin: 0 20px;
		cursor: pointer;
	}

</style>
