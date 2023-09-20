<template>
  <div class="home">
    <el-button type="primary" @click="exportExcel">设计表格</el-button>
    <div>
      参考：https://www.jianshu.com/p/1f858b9cdaf3
    </div>
    <div>
      参考：https://blog.csdn.net/ljy123w/article/details/124005869
    </div>
    <div>
      预览表格：https://github.com/501351981/vue-office
    </div>
    <xExcel :drawer.sync="drawer"></xExcel>
  </div>
</template>

<script>
import xExcel from "../components/xExcel/xExcel";
export default {
  name: 'Home',
  components: {
    xExcel
  },
  data() {
    return {
      drawer: false,
    }
  },
  methods: {
    exportExcel(){
      this.drawer = true
    }
  }
}
</script>


