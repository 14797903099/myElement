<template>
  <div>
      我是  b页
      <el-button @click="handlePort">b导出</el-button>
  </div>
</template>

<script>
	
export default {
  name: 'hello',
  data () {
    return {
      msg: 'Welcome to Your Vue.js App'
    }
  },
  methods:{
    handlePort(){
      // let val = ['MachineNo', 'MachineIndex', 'ProdMonth', 'MachineType', 'StatusStr'];
      //           let proArr = ['StartTime', 'EndTime', 'ValidTime'];
      //           let outArr = ['CaseNo', 'ToCountry', 'ToPlace', 'OutTime', 'InvertoryNum', 'OrderTime'];
      //           let tsArr = ['InstallPlace', 'InstallTime', 'CompCount', 'FinCompTime'];
                // val = val.concat(opColKey).concat(acColKey).concat(proArr).concat(partNameArr).concat(outArr).concat(tsArr);
                // let startCell = 5 + opColName.length + acColName.length;//以下是合并单元格
                //  let startCell =  6
                // let merges = [{
                //         //合并基本情报
                //         s: { //s为开始
                //             c: 0, //开始列
                //             r: 0 //开始取值范围
                //         },
                //         e: { //e结束
                //             c: (startCell - 1), //结束列
                //             r: 0 //结束范围
                //         }
                //     },
                //     {
                //         //合并生产情报
                //         s: { //s为开始
                //             c: startCell, //开始列
                //             r: 0 //开始取值范围
                //         },
                //         e: { //e结束
                //             c: (startCell + 3 - 1), //结束列
                //             r: 0 //结束范围
                //         }
                //     },
                //     {
                //         //合并部件情报
                //         s: { //s为开始
                //             c: (startCell + 3), //开始列
                //             r: 0 //开始取值范围
                //         },
                //         e: { //e结束
                //             c: (startCell + 3 + this.partNameHeader.length - 1), //结束列
                //             r: 0 //结束范围
                //         }
                //     },
                //     {
                //         //合并出库情报
                //         s: { //s为开始
                //             c: (startCell + 3 + this.partNameHeader.length), //开始列
                //             r: 0 //开始取值范围
                //         },
                //         e: { //e结束
                //             c: (startCell + 3 + this.partNameHeader.length + 6 - 1), //结束列
                //             r: 0 //结束范围
                //         }
                //     },
                //     {
                //         //合并技服情报
                //         s: { //s为开始
                //             c: (startCell + 3 + this.partNameHeader.length + 6), //开始列
                //             r: 0 //开始取值范围
                //         },
                //         e: { //e结束
                //             c: (startCell + 3 + this.partNameHeader.length + 6 + 4 - 1), //结束列
                //             r: 0 //结束范围
                //         }
                //     },
                // ];
               
//                  let tHeader = [
// 		'鲜花',
// 		'颜色',
//     '照片']
//              let data = [
// 		{
// 			name: '玫瑰花',
// 			color: '红色',
// 			pic: 'https://ss1.bdstatic.com/70cFuXSh_Q1YnxGkpoWK1HF6hhy/it/u=2801998497,4036145562&fm=27&gp=0.jpg'
// 		},
// 		{
// 			name: '菊花',
// 			color: '黄色',
// 			pic: 'https://ss1.bdstatic.com/70cFuXSh_Q1YnxGkpoWK1HF6hhy/it/u=1506844670,1837003941&fm=200&gp=0.jpg'
// 		},
// 		{
// 			name: '牵牛花',
// 			color: '紫色',
// 			pic: 'https://ss1.bdstatic.com/70cFuXSh_Q1YnxGkpoWK1HF6hhy/it/u=3056120770,1115785765&fm=27&gp=0.jpg'
// 		},
// 		{
// 			name: '梅花',
// 			color: '白色',
// 			pic: 'https://ss1.bdstatic.com/70cFvXSh_Q1YnxGkpoWK1HF6hhy/it/u=2700343322,3431874915&fm=27&gp=0.jpg'
// 		},
// 		{
// 			name: '桃花花',
// 			color: '粉色',
// 			pic: 'https://ss1.bdstatic.com/70cFvXSh_Q1YnxGkpoWK1HF6hhy/it/u=602076004,4209938077&fm=27&gp=0.jpg'
// 		}
//              ]
//              let multiHeader  = []
let merges = [
  // {
  //     s: { //s为开始
  //         c: 2, //开始列
  //         r: 0 //开始取值范围
  //     },
  //     e: { //e结束
  //         c: 3, //结束列
  //         r: 0 //结束范围
  //     }
  // }
]  
// let val = []
    const multiHeader = [['牛奶', '果汁', '食品数量', null, '零食数量', null, null, '蔬菜数量', null, null]] //对应表格输出的title
        const tHeader = [[null, null, '甜甜圈', '干脆面', '面包', '山药薄片', '薯片', '菠菜', '芹菜', '大油菜']] //对应表格输出的title
            const val = ['milk', 'juice', 'ttq', 'noodle', 'break', 'crispss', 'chips', 'spinach', 'celery', 'Canola'] // 对应表格tableData输出的数据
    const data = [
      {
        milk: '1',
        juice: '1',
        ttq: '1',
        noodle: '1',
        break: '1',
        crispss: '1',
        chips:'1',
        spinach:'1',
        celery: '1',
        Canola: '1'
      },
         {
        milk: '1',
        juice: '1',
        ttq: '1',
        noodle: '1',
        break: '1',
        crispss: '1',
        chips:'1',
        spinach:'5',
        celery: '4',
        Canola: '3'
      }
    ]
    //  this.exportExcel(multiHeader, tHeader, merges, val, data);

    },
    // exportExcel(multiHeader, tHeader, merges, val, tableData) {
    //       require.ensure([], () => {
    //           const {
    //               export_json_to_excel
    //           } = require('../../../static/excel/Export2Excel.js');
    //           let filterVal = val;
    //           //把后台传过来的数据存到list
    //           const list = tableData;
    //           const data = this.formatJson(filterVal, list);
    //           export_json_to_excel(multiHeader, tHeader, merges, data, '统合情报管理报表');
    //       })
    //   },
      formatJson(filterVal, jsonData) {
          return jsonData.map(v => filterVal.map(j => v[j]));
      }
  } 
}
</script>
	
<!-- Add "scoped" attribute to limit CSS to this component only -->
<style scoped>
.hello{
	width: 100%;
	height: 100%;
	background: #ccc;
}
</style>
