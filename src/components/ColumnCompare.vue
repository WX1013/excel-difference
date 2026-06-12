<template>
  <div class="page-wrap">
    <a-card class="main-card">
      <a-spin :spinning="processing" tip="处理中...">
        <a-space direction="vertical" :size="24" style="width: 100%">
          <a-typography-title :level="4" class="page-title">
            第一列与第二列文本对比
          </a-typography-title>

          <a-row>
            <a-col :span="24">
              <div class="desc-note">
                <p>读取 Excel 第一个 Sheet 的 <strong>第一列</strong> 和 <strong>第二列</strong>，逐行进行文本比对。</p>
                <p>两列内容一致则在 <strong>第三列</strong> 标记 <strong>1</strong>，不一致则标记 <strong>0</strong>。</p>
              </div>
            </a-col>
          </a-row>

          <a-upload-dragger
            accept=".xlsx,.xls"
            :max-count="1"
            :before-upload="beforeUpload"
            @change="onUploadChange"
            :show-upload-list="false"
            :disabled="processing"
          >
            <p class="ant-upload-drag-icon">
              <UploadOutlined />
            </p>
            <p class="ant-upload-text">点击上传或拖入上传区</p>
            <p class="ant-upload-hint">仅支持 Excel 文件（.xlsx / .xls）</p>
            <p v-if="selectedFile" class="ant-upload-file-name">{{ selectedFile.name }}</p>
          </a-upload-dragger>

          <div class="action-area">
            <a-button
              type="primary"
              size="large"
              :disabled="!selectedFile || processing"
              :loading="processing"
              @click="onStartProcess"
            >
              {{ processing ? '处理中...' : '开始对比并下载' }}
            </a-button>
          </div>

          <a-alert
            v-if="statusMessage"
            :message="statusMessage"
            :type="statusType"
            show-icon
            class="status-alert"
          />
        </a-space>
      </a-spin>
    </a-card>
  </div>
</template>

<script setup lang="ts">
import { ref, computed } from 'vue';
import { UploadOutlined } from '@ant-design/icons-vue';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

const selectedFile = ref<File | null>(null);
const processing = ref(false);
const statusMessage = ref('');

const statusType = computed(() => {
  if (!statusMessage.value) return 'info';
  if (statusMessage.value.startsWith('处理失败')) return 'error';
  if (statusMessage.value.includes('已下载')) return 'success';
  return 'info';
});

function beforeUpload() {
  return false;
}

function onUploadChange(info: any) {
  const file = info?.file?.originFileObj ?? info?.file ?? null;
  selectedFile.value = file;
  statusMessage.value = file ? `已选择文件：${file.name}` : '';
}

async function onStartProcess() {
  if (!selectedFile.value) return;

  processing.value = true;
  statusMessage.value = '正在对比...';

  try {
    const arrayBuffer = await selectedFile.value.arrayBuffer();
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(arrayBuffer);

    const worksheet = workbook.worksheets[0];
    if (!worksheet) {
      throw new Error('Excel 中未找到任何工作表');
    }

    // 写入表头
    worksheet.getCell(1, 3).value = '对比结果';

    let processedCount = 0;
    let matchCount = 0;
    const lastRowNumber = worksheet.lastRow?.number ?? worksheet.rowCount;

    for (let r = 2; r <= lastRowNumber; r++) {
      const colA = String(worksheet.getCell(r, 1).value ?? '').trim();
      const colB = String(worksheet.getCell(r, 2).value ?? '').trim();

      if (!colA && !colB) continue;

      const result = colA === colB ? 1 : 0;
      worksheet.getCell(r, 3).value = result;
      processedCount++;
      if (result === 1) matchCount++;
    }

    const outBuffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([outBuffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    const originalName = selectedFile.value.name.replace(/\.xlsx?$/i, '');
    const outName = `${originalName}_列对比.xlsx`;
    saveAs(blob, outName);

    statusMessage.value = `对比完成，共 ${processedCount} 行：一致 ${matchCount} 行 / 不一致 ${processedCount - matchCount} 行，已下载结果文件。`;
  } catch (err: any) {
    console.error(err);
    statusMessage.value = `处理失败：${err?.message || '未知错误'}`;
  } finally {
    processing.value = false;
  }
}
</script>

<style scoped>
.page-wrap {
  min-height: 100vh;
  padding: 24px;
  background: linear-gradient(135deg, #f5f7fa 0%, #e4e8ec 100%);
  display: flex;
  align-items: center;
  justify-content: center;
}
.main-card {
  max-width: 560px;
  width: 100%;
  box-shadow: 0 2px 12px rgba(0, 0, 0, 0.08);
}
.page-title {
  margin-bottom: 0 !important;
  text-align: center;
}
.desc-note {
  background: #e6f7ff;
  padding: 10px 12px;
  border-radius: 6px;
  border: 1px solid rgba(24, 144, 255, 0.12);
  font-size: 12px;
  color: rgba(0, 0, 0, 0.75);
  line-height: 1.6;
}
.desc-note p {
  margin: 4px 0;
}
.action-area {
  display: flex;
  justify-content: center;
}
.ant-upload-file-name {
  margin-top: 8px;
  color: var(--ant-colorPrimary);
  font-size: 13px;
}
.status-alert {
  margin-top: 8px;
}
</style>
