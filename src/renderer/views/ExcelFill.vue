<template>
  <div>
    <a-breadcrumb style="margin: 16px 0">
      <a-breadcrumb-item>Excel</a-breadcrumb-item>
      <a-breadcrumb-item>多表数据填充</a-breadcrumb-item>
    </a-breadcrumb>
    <div
      :style="{ padding: '24px', background: '#fff', minHeight: '360px' }"
    >
        <c-table></c-table>
    </div>
  </div>
</template>
<script>
import Table from "@/components/Fill";
export default {
  components: {
    "c-table": Table,
  },
};
</script>