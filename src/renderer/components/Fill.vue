<template>
  <div>
    <table class="table">
      <thead>
        <tr id="excel-th">
          <th>文件名</th>
          <th v-for="(row, index) in rows" :key="index">
            匹配项{{ index + 1 }}
          </th>
          <th>需填充的项</th>
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
            <a-select style="width: 120px">
              <a-select-option
                v-for="(key, index_fk) in title.get('keyTitle')"
                :value="key"
                :key="index_fk"
                @click="addFillKey(index, key)"
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
    <a-button type="default" @click="addRow"> 增加匹配行（列） </a-button>
    <a-button type="default" @click="delRow">删除匹配行（列）</a-button>
    <a-button type="primary" @click="run" icon="download">开始匹配</a-button>
  </div>
</template>
<script>
import XLSX from "xlsx";
import { dealData, save } from "@/assets/js/sheet";

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
        var rowObject = XLSX.utils.sheet_to_json(worksheet, {defval: ""});
        temp.set("keyTitle", Object.keys(rowObject[0]));
        temp.set("data", rowObject);
        _this.titles.push(temp);
      };
      reader.readAsBinaryString(files[0]);
    },
    delFile(index) {
      this.titles.splice(index, 1);
    },
    addFillKey(index, key) {
      this.titles[index].set("fillkey", key);
    },
    addKey(fileIndex, rowIndex, value) {
      this.titles[fileIndex].set("match" + rowIndex, value);
    },
    async run() {
      const [cacheMatchValue, cacheMatchedSheetList] = dealData(
        this.titles,
        this.rows
      );
      let filledkey = this.titles[0].get("fillkey");
      for (let matchIndex in cacheMatchValue) {
        for (let cacheMatchedSheetIndex in cacheMatchedSheetList) {
          let cacheMatchedSheet = cacheMatchedSheetList[cacheMatchedSheetIndex];
          let index = cacheMatchedSheet.indexOf(cacheMatchValue[matchIndex]);
          cacheMatchedSheetIndex++;
          let currentSheet = this.titles[cacheMatchedSheetIndex];
          let fillkey = currentSheet.get("fillkey");
          if (index !== -1) {
            let value = currentSheet.get("data")[index][fillkey];
            this.titles[0].get("data")[matchIndex][filledkey] = value;
            break;
          }
        }
      }
      await save(this.titles[0].get("data"));
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