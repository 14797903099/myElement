<template>
  <div>
    <el-button @click="handlePort">导出</el-button>
  </div>
</template>

<script>
import XLSX from 'xlsx'
import {sheet2blob, openDownloadDialog } from '../../../static/js/exportExcel'
export default {
  name: 'hello',
  data () {
    return {
      msg: 'Welcome to Your Vue.js App'
    }
  },
  methods:{
handlePort(){
  	var aoa = [
		[null, null, null, '其它信息'], // 特别注意合并的地方后面预留2个null
		['姓名', '性别', '年龄', '注册时间'],
		[null, '男', 18, new Date()],
		['李四', '女', 22, new Date()]
	];
	var sheet = XLSX.utils.aoa_to_sheet(aoa);
	sheet['!merges'] = [
		// 设置A1-C1的单元格合并
		// {	s: {r: 0, c: 0}, 
		// 	e: {r: 0, c: 2}
		// },
		// {	s: {r: 2, c: 0}, 
		// 	e: {r: 2, c: 3}
		// }
	];
	openDownloadDialog(sheet2blob(sheet), '单元格合并示例.xlsx');
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
