<template>
  <div class="app-container">
    <a-tabs v-model:activeKey="activeTab" centered size="large">
      <a-tab-pane key="column-split" tab="分列">
        <ColumnSplit />
      </a-tab-pane>
      <a-tab-pane key="column-compare" tab="列对比">
        <ColumnCompare />
      </a-tab-pane>

    </a-tabs>
  </div>
</template>

<script setup lang="ts">
import { ref } from 'vue';

import ColumnSplit from './components/ColumnSplit.vue';
import ColumnCompare from './components/ColumnCompare.vue';

const activeTab = ref('column-split');
</script>

<style scoped>
.app-container {
  min-height: 100vh;
  background: linear-gradient(135deg, #f5f7fa 0%, #e4e8ec 100%);
}
</style>
