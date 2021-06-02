<template>
  <div>
    <table class="table">
      <thead>
        <tr id="excel-th">
          <th>文件名</th>
          <th v-for="(row, index) in rows" :key="index">
            匹配行（列）{{ index + 1 }}
          </th>
        </tr>
      </thead>
      <tbody>
        <tr v-for="(title, index) in titles" :key="index">
          <td>{{ title.get("name") }}</td>
          <td v-for="(row, index_r) in rows" :key="index_r">
            <a-select style="width: 120px">
              <a-select-option
                v-for="(key, index_k) in title.get('keyTitle')"
                :value="key"
                :key="index_k"
                @click="addKey(index, index_r, key)"
                >{{ key }}</a-select-option
              >
            </a-select>
          </td>
          <td>
            <a-button
              type="danger"
              shape="circle"
              icon="close"
              @click="delFile(index)"
            ></a-button>
          </td>
        </tr>
      </tbody>
    </table>
    <input type="file" id="excel" @change="readXLSX($event)" />
      <a-button type="default" @click="addRow">
      增加匹配行（列）
    </a-button>
    <a-button type="default" @click="delRow">删除匹配行（列）</a-button>
    <a-button type="primary" @click="run" icon="download">开始匹配</a-button>
  </div>
</template>
<script>
import XLSX from "xlsx";
const { dialog } = require("electron").remote;

export default {
  data() {
    return {
      rows: 2,
      titles: [],
      data: "",
      title: new Map(),
    };
  },
  methods: {
    addRow() {
      this.rows++;
    },
    delRow() {
      if (this.rows >= 2) {
        this.rows--;
      }
    },
    handleChange(value, index) {
      console.log(`selected ${value}`);
    },
    readXLSX(e) {
      let _this = this;
      let files = e.target.files;
      let temp = new Map();
      temp.set("name", files[0].name);
      let reader = new FileReader();
      reader.onload = function (e) {
        let data = e.target.result;
        let workbook = XLSX.read(data, { type: "binary" });
        let sheetName = workbook.SheetNames[0];
        let worksheet = workbook.Sheets[sheetName];
        var rowObject = XLSX.utils.sheet_to_row_object_array(worksheet);
        temp.set("keyTitle", Object.keys(rowObject[0]));
        temp.set("data", rowObject);
        _this.titles.push(temp);
      };
      reader.readAsBinaryString(files[0]);
    },
    delFile(index) {
      this.titles.splice(index, 1);
    },
    addKey(fileIndex, rowIndex, value) {
      this.titles[fileIndex].set("match" + rowIndex, value);
    },
    async run() {
      let cacheTable = [];
      let matchFile = this.titles[0];
      let matchKeys = [];
      let matchedKeys = [];
      for (let i = 0; i < this.rows; i++) {
        matchKeys.push(matchFile.get("match" + i));
        matchedKeys.push(this.titles[1].get("match" + i));
      }
      // 处理matchFile 数据
      let cacheMatchValue = [];
      for (let valueMatchFile of matchFile.get("data")) {
        let cache = [];
        let cacheStr = "";
        for (let matchKey of matchKeys) {
          console.log(valueMatchFile)
          cache.push(valueMatchFile[matchKey].trim());
          // cache[matchIndex] =  valueMatchFile[matchKeys[matchIndex]]
        }
        cacheStr = JSON.stringify(cache);
        cacheMatchValue.push(cacheStr);
      }
      // 处理matchedFile 数据
      let cacheMatchedValue = [];
      for (let valueMatchedFile of this.titles[1].get("data")) {
        let cache = [];
        let cacheStr = "";
        for (let matchedKey of matchedKeys) {
          cache.push(valueMatchedFile[matchedKey]);
        }
        cacheStr = JSON.stringify(cache);
        cacheMatchedValue.push(cacheStr);
      }
      // 开始匹配
      for (let matchIndex in cacheMatchValue) {
        let index = cacheMatchedValue.indexOf(cacheMatchValue[matchIndex]);
        if (index !== -1) {
          cacheTable.push(this.titles[1].get("data")[index]);
        }
      }
      console.log(cacheTable);
      let ws = XLSX.utils.json_to_sheet(cacheTable);
      let wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "test");
      console.log(dialog);
      let filePath = await dialog.showSaveDialog({
        defaultPath: "generate.xlsx",
      });
      if (filePath) {
        XLSX.writeFile(wb, filePath);
      }
    },
  },
};
</script>
<style scoped>
.table {
  width: 100%;
  margin-bottom: 30px;
}
</style>