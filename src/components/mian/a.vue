<template>
  <div>
     我是啊
  </div>
</template>

<script>
	
export default {
  name: 'hello',
  data () {
    return {
      msg: 'Welcome to Your Vue.js App'
    }
  }
}
</script>
	
<!-- Add "scoped" attribute to limit CSS to this component only -->
<style lang="less">
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
