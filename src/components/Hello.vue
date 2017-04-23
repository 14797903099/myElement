<template>
  <div class="hello">
      我是  hello页
      
  </div>
</template>

<script>

export default {
  name: 'hello',
  data () {
    return {
    }
  },
  created:function(){
	console.log(ajax)
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
