<template>
  <div>
     <!--<div :class="{'class-a':isA, 'class-b':!isA}" @click="handleTab">我是三角阀技术开发及</div>-->
     <ul>
       <li v-for='item in data' v-text="item.name" :class="{cur:item.iscur}" @click="setCur($index)"></li>
     </ul> 
  </div>
</template>

<script>
	
export default {
  name: 'hello',
  data () {
    return {
      msg: 'Welcome to Your Vue.js App',
      isA:false
    }
  },
  methods: {
    handleTab () {
      this.isA = !this.isA
    },
    setCur (index) {
      this.data.map(function (v,i) {
          i==index? v.iscur=true: v.iscur=false;
      });
    } 
  }
}
</script>
	
<!-- Add "scoped" attribute to limit CSS to this component only -->
<style lang="less">
.class-a{
  color: red
}
.class-b{
  color: green
}
.hello{
	width: 100%;
	height: 100%;
	background: #ccc;
}
 .el-carousel__item h3 {
    color: #475669;
    font-size: 14px;
    opacity: 0.75;
    line-height: 150px;
    margin: 0;
  }

  .el-carousel__item:nth-child(2n) {
     background-color: #99a9bf;
  }
  
  .el-carousel__item:nth-child(2n+1) {
     background-color: #d3dce6;
  }
</style>
