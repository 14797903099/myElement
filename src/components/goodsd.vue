<template>
  <div class="header" @click='click'>
  	fjskfjksjdfkasjkfdjasf
  	<span>{{text}}</span>
  	<h2 :id="index" @click="test">fsdfsfdsdf</h2>
  	<h3 ref="h" @click="tt">fkdfjkasf</h3>
  </div>
</template>

<script>
		import {ajax} from '@/ajax.js';
export default {
	data () {
    return {
//    message:"张三",
      index:1,
      text:"niho"
    }
  },
  methods: {
  	tt(){
  		console.log(this.$refs.h);
  	},
  	test(e){
//		console.log(e.target.id);
  	},
  	click(e) {
  	
  		var that =this;
  		this.$emit("mess",JSON.stringify(["张三","李四","王五"]));
//		ajax({
//			url:"http://localhost:8080/mock/data.json",
//			success:function(res){
//				console.log(JSON.parse(res));
//				that.msg=JSON.parse(res);
//			}
//		})
			this.$http.get('http://localhost:8080/mock/data.json')
                    .then((response) => { //异步
                        if (response) {
                            console.log(response.data.goods);
                        } 
                    });
  	}
  }
}
</script>
	
<style lang="less">
.header{
		height:100px;
		background:green;
}
</style>
