<template>
  <el-drawer
      title="导出设计"
      :visible.sync="show"
      size="90%"
      :direction="direction">
    <div class="office-body">
      <div class="office-preview" ref="x-spreadsheet-demo">
        <div class="spreadsheet" v-if="drawer" id="x-spreadsheet-demo"></div>
      </div>
      <div class="office-design">
        <el-button type="primary" @click="exportHandler">导出表格</el-button>
        <el-button type="primary" @click="exportConfHandler">导出配置</el-button>
      </div>
    </div>
  </el-drawer>
</template>

<script>
import * as XLSX from 'xlsx'
import XLSXD from './modules/xlsx-style'
import Spreadsheet from "x-data-spreadsheet";
import 'x-data-spreadsheet/dist/locale/zh-cn';
import _ from "lodash";
Spreadsheet.locale('zh-cn')

export default {
  name: "xExcel",
  props:{
    drawer: {
      type: Boolean,
      default: false
    }
  },
  data(){
    return {
      direction: 'rtl',
      s: null,
      alphabet: [], // 字母表
      sheetConf: {
        mode: 'edit', // edit | read
        showToolbar: true,
        showGrid: true,
        showContextmenu: true,
        view: {
          height: () =>
              this.$refs['x-spreadsheet-demo'] &&
              this.$refs['x-spreadsheet-demo'].offsetHeight &&
              _.isNumber(this.$refs['x-spreadsheet-demo'].offsetHeight)
                  ? this.$refs['x-spreadsheet-demo'].offsetHeight
                  : 0,
          width: () =>
              this.$refs['x-spreadsheet-demo'] &&
              this.$refs['x-spreadsheet-demo'].offsetWidth &&
              _.isNumber(this.$refs['x-spreadsheet-demo'].offsetWidth)
                  ? this.$refs['x-spreadsheet-demo'].offsetWidth
                  : 0,
        },
        row: {
          len: 100,
          height: 25,
        },
        col: {
          len: 26,
          width: 80,
          indexWidth: 60,
          minWidth: 60,
        },
        style: {
          bgcolor: '#ffffff',
          align: 'left',
          valign: 'bottom',
          textwrap: false,
          strike: false,
          underline: false,
          color: '#0a0a0a',
          font: {
            name: 'Helvetica',
            size: 10,
            bold: false,
            italic: false,
          },
        },
      },
    }
  },
  computed:{
    show: {
      get(){
        if(this.drawer){
          this.init()
        }
        return this.drawer
      },
      set(val){
        this.$emit('update:drawer', val)
      }
    }
  },
  mounted() {
    this.alphabet = this.alphabet26()
  },
  methods:{
    init(){

      this.$nextTick(() => {
        this.s = new Spreadsheet("#x-spreadsheet-demo", this.sheetConf)
            .loadData({}) // load data
            .change(data => {
              // 数据改变
              console.log(data)
            });
        this.s.validate()
      })
    },
    exportHandler(){
      let wb = this.xtos(this.s.getData())
      let tmpDown = new Blob([
        this.s2ab(
            XLSXD.write(wb, {
              bookType: "xlsx",
              bookSST: true,
              type: "binary",
              cellStyles: true,
            })
        ),
      ]);
      this.downExcel(tmpDown, "测试.xlsx");
    },
    exportConfHandler(){
      let wb = this.xtos(this.s.getData())
      // 导出JSON
      let jsonObj = JSON.stringify(wb)
      // jsonObj 转文件流
      let tmpDown = new Blob([jsonObj]);
      this.downExcel(tmpDown, "测试.json");
    },
    xtos(sdata){
      let sheetNames = [];
      let sheetArr = {}
      const wb = XLSX.utils.book_new();
      sdata.forEach((xws) => {
        let rowobj = xws.rows;
        let styleArr = xws.styles;
        sheetNames.push(xws.name);
        // 读取列宽
        let cols = new Array(xws.cols.len);
        // 行高
        let rows = new Array(xws.rows.len);
        for (let col in xws.cols) {
          if (xws.cols[col].width) {
            // 指定位置替换
            cols.splice(Number(col), 1, {
              wch: xws.cols[col].width / 8,
            });
          }
        }
        for (let row in xws.rows) {
          if (xws.rows[row].height) {
            // 指定位置替换
            rows.splice(Number(row), 1, {
              hpx: xws.rows[row].height - this.sheetConf.row.height,
            });
          }
        }
        if(!this.isObject(cols)) cols = [];
        if(!this.isObject(rows)) rows = [];
        let aoa = [[]];
        let merges = [];
        let alphabetBorder = {};
        for(let ri = 0; ri < rowobj.len; ++ri) {
          let row = rowobj[ri];
          if(!row) continue;
          aoa[ri] = [];
          Object.keys(row.cells).forEach((k) => {
            let idx = +k;
            if(isNaN(idx)) return;
            aoa[ri][idx] = row.cells[k].text;

            // 合并单元格
            if (row.cells[k].merge) {
              // 开始行
              let startRow = ri;
              // 结束行,加上Y轴跨度
              let endRow = startRow + row.cells[k].merge[0];
              // 开始列
              let startColumn = Number(k);
              // 结束列,加上X轴跨度
              let endColumn = startColumn + row.cells[k].merge[1];
              // 按开始行，开始列，结束行，结束列合并
              merges.push({s: {r: startRow, c: startColumn}, e: {r: endRow, c: endColumn}});
            }
            // 单元格样式
            if (row.cells[k].style !== undefined && row.cells[k]?.style !== null) {
              let rowsIndex = ri
              // 列的索引 k
              // console.log(rowsIndex, '行索引')
              // console.log(k, '列索引')
              // console.log(row.cells[k].style, '样式索引')
              // 新增
              let styleRow = rowsIndex + 1
              let alphabet = this.alphabet[k]
              alphabetBorder[alphabet+styleRow] = {s:{
                  border: {},
                  alignment: {
                    horizontal: "left", // 水平居中
                    vertical: "bottom", // 垂直居中
                  },
                  font: {
                    name: "微软雅黑",
                    sz: this.sheetConf.style.font.size || 10,
                    bold: false,
                    color: {},
                  },
              }}
              // 边框样式
              if(styleArr[row.cells[k].style]?.border){
                Object.keys(styleArr[row.cells[k].style]?.border).forEach((key) => {
                  alphabetBorder[alphabet+styleRow].s.border[key] = {
                    style: styleArr[row.cells[k].style].border[key][0],
                    color: {
                      rgb: this.toRgb(styleArr[row.cells[k].style].border[key][1])
                      // auto: 1
                    }
                  }
                })
              }
              // 文字垂直样式
              if(styleArr[row.cells[k].style]?.valign){
                alphabetBorder[alphabet+styleRow].s.alignment.vertical = styleArr[row.cells[k].style].valign === 'middle' ? 'center' : styleArr[row.cells[k].style].valign
              }
              // 文字水平样式
              if(styleArr[row.cells[k].style]?.align){
                alphabetBorder[alphabet+styleRow].s.alignment.horizontal = styleArr[row.cells[k].style].align === 'middle' ? 'center' : styleArr[row.cells[k].style].align
              }
              // 文字颜色
              if(styleArr[row.cells[k].style]?.color){
                alphabetBorder[alphabet+styleRow].s.font.color['rgb'] = this.toRgb(styleArr[row.cells[k].style].color)
              }
              // 背景色
              if(styleArr[row.cells[k].style]?.bgcolor){
                alphabetBorder[alphabet+styleRow].s['fill'] = {
                  bgColor: {indexed: 64},
                  fgColor: {
                    rgb: this.toRgb(styleArr[row.cells[k].style].bgcolor)
                  }
                }
              }
              // 文字是否加粗
              if(styleArr[row.cells[k].style].font?.bold !== undefined || styleArr[row.cells[k].style].font?.bold !== null){
                alphabetBorder[alphabet+styleRow].s.font.bold = styleArr[row.cells[k].style]?.font?.bold
              }
            }
          });
        }
        const ws = XLSX.utils.aoa_to_sheet(aoa)

        // ws对象与单元格样式合并
        for (let wsKey in ws) {
          if(wsKey === '!ref') continue
          if(alphabetBorder[wsKey]){
            ws[wsKey].s = alphabetBorder[wsKey].s
          }
          // 默认字体样式
          if(!ws[wsKey]['s']){
            ws[wsKey]['s'] = {}
            ws[wsKey]['s']['font'] = {
              name: "微软雅黑",
              sz: this.sheetConf.style.font.size || 10,
              bold: false,
              color: {},
            }
          }
        }
        sheetArr[xws.name] = {
          ...ws,
          ['!cols']: cols, // 列宽
          ['!rows']: rows, // 行高
          ['!merges']: merges // 单元格合并
        };
      })
      sheetNames.forEach(item => {
        XLSX.utils.book_append_sheet(wb, sheetArr[item], item)
      })
      return wb;
    },
    // 自定义下载文件方法
    downExcel(obj, fileName) {
      const a_node = document.createElement("a");
      a_node.download = fileName;
      if ("msSaveOrOpenBlob" in navigator) {
        window.navigator.msSaveOrOpenBlob(obj, fileName);
      } else {
        a_node.href = URL.createObjectURL(obj);
      }
      a_node.click();
      //
      setTimeout(() => {
        URL.revokeObjectURL(obj);
      }, 2000);
    },
    // 文件流转换
    s2ab(s) {
      if (typeof ArrayBuffer !== "undefined") {
        const buf = new ArrayBuffer(s.length);
        const view = new Uint8Array(buf);
        for (let i = 0; i != s.length; ++i) {
          view[i] = s.charCodeAt(i) & 0xff;
        }
        return buf;
      } else {
        const buf = new Array(s.length);
        for (let i = 0; i != s.length; ++i) {
          buf[i] = s.charCodeAt(i) & 0xff;
        }
        return buf;
      }
    },
    // 26个大写英文字母 65大写 97小写
    alphabet26() {
      let arr = [];
      for (let i = 0; i < 26; i++) {
        arr.push(String.fromCharCode(65 + i));
      }
      return arr;
    },
    // 处理颜色
    toRgb(color){
      // 干掉＃
      let fullColor = /^#/.test(color) ? color.replace('#', '') : color
      return fullColor
    },
    // 判断数组中是否包含对象类型
    isObject(arr){
      let flag = false
      let len = arr.length
      for (let i = 0; i < len; i++) {
        if(typeof arr[i] === 'object'){
          flag = true
          break
        }
      }
      return flag
    }
  }
}
</script>

<style lang="scss" scoped>
.office-body{
  display: flex;
  height: 100%;
  .office-preview{
    flex: 1;
    .spreadsheet{
      width: 100%;
      height: 100%;
    }
  }
  .office-design{
    flex-shrink: 0;
    width: 30%;
    background-color: #42b983;
  }
}
</style>
