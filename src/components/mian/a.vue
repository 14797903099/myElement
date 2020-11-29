<template>
  <div>
    <el-button @click="handlePort">导出</el-button>
     <el-button @click="handleGet">获取参数</el-button>
     <el-table
      :data="tableData"
       :span-method="objectSpanMethod"
      border
    style="width: 100%">
    <el-table-column
      :prop="multiHeader[0].prop"
      :label="multiHeader[0].label"
      width="150">
    </el-table-column>
    <el-table-column
     :prop="multiHeader[1].prop"
      :label="multiHeader[1].label"
      width="120">
    </el-table-column>
    <el-table-column :label="multiHeader[2].label">
      <template v-for="(col, i) in multiHeader[2].tt">
        <el-table-column
          :prop="col.prop"
         :label="col.label"
         :key="i"
          width="120">
        </el-table-column>
      </template>
    </el-table-column>
  </el-table>
  </div>
</template>

<script>
import XLSX from 'xlsx'
import {sheet2blob, openDownloadDialog } from '../../../static/js/exportExcel'
export default {
  name: 'hello',
  data () {
    return {
      msg: 'Welcome to Your Vue.js App',
       multiHeader:[
         {label:'日期', prop:'date'}, 
         {label:'姓名', prop:'name'},
         {
          label:'地址',
          tt:[
          {label:'省份', prop:'province'}, 
          {label:'市区', prop:'city'},
          {label:'地址', prop:'address'},
          {label:'邮编', prop:'zip'}
          ]
         }
        ],
       tableData: [],
        spanArr: [],
        pos: 0,
        list: [{
         date: '2016-05-03',
          name: '王小虎',
          province: '上海',
          city: '普陀区',
          address: '上海市普陀区金沙江路 1518 弄',
          zip: 200333
        }, {
          date: '2016-05-03',
          name: '王小虎',
          province: '上海',
          city: '普陀区',
          address: '上海市普陀区金沙江路 1518 弄',
          zip: 200333
        }, {
          date: '2016-05-03',
          name: '王小虎',
          province: '上海',
          city: '普陀区',
          address: '上海市普陀区金沙江路 1518 弄',
          zip: 200333
        }, {
          date: '2016-05-01',
          name: '王小虎',
          province: '上海',
          city: '普陀区',
          address: '上海市普陀区金沙江路 1518 弄',
          zip: 200333
        },
        {
          date: '2016-05-01',
          name: '王小虎',
          province: '上海',
          city: '普陀区',
          address: '上海市普陀区金沙江路 1518 弄',
          zip: 200333
        },
        {
          date: '2016-05-01',
          name: '王小虎',
          province: '上海',
          city: '普陀区',
          address: '上海市普陀区金沙江路 1518 弄',
          zip: 200333
        }]
    }
  },
  methods:{
     objectSpanMethod({ row, column, rowIndex, columnIndex }) {
        if (columnIndex === 0) {
          console.log(1)
        const _row = this.spanArr[rowIndex];
        const _col = _row > 0 ? 1 : 0;
        return {
          rowspan: _row,
          colspan: _col
        };
      }
      },
      handleGet () {
         let scenceTypeList = this.list.map(e => {
            return e.date;
          });
          this.getSpanArr(scenceTypeList)
      },
      getSpanArr(data) {
      for (var i = 0; i < data.length; i++) {
        if (i === 0) {
          this.spanArr.push(1);
          this.pos = 0;
        } else {
          // 判断当前元素与上一个元素是否相同,因合并第一个所以[0]
          if (data[i] === data[i - 1]) {
            this.spanArr[this.pos] += 1;
            this.spanArr.push(0);
          } else {
            this.spanArr.push(1);
            this.pos = i;
          }
        }
      }
      console.log(this.spanArr);
      this.tableData = this.list
    },
     formatJson(filterVal, jsonData) {
          return jsonData.map(v => filterVal.map(j => v[j]));
      },
handlePort(){
  // 	var aoa = [
	// 	[null, null, null, '其它信息'], // 特别注意合并的地方后面预留2个null
	// 	['姓名', '性别', '年龄', '注册时间'],
	// 	[null, '男', 18, new Date()],
	// 	['李四', '女', 22, new Date()]
  // ];   
      let filterVal = ['date','province', 'city', 'address']
      let aoa = this.formatJson(filterVal, this.list)
      aoa.forEach((item, index) => {
        if (this.spanArr[index] == 0){
          item[0] = null
        }
      });
      console.log(aoa)
  // 	var aoa = [
	// 	[null, null, null, '其它信息'], // 特别注意合并的地方后面预留2个null
	// 	['姓名', '性别', '年龄', '注册时间'],
	// 	[null, '男', 18, new Date()],
	// 	['李四', '女', 22, new Date()]
	// ];
	var sheet = XLSX.utils.aoa_to_sheet(aoa);
	sheet['!merges'] = [
		// 设置A1-C1的单元格合并
		{	s: {r: 0, c: 0}, 
			e: {r: 2, c: 0}
		},
		{	s: {r: 3, c: 0}, 
			e: {r: 5, c: 0}
		}
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
