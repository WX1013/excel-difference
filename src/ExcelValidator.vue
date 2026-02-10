<template>
  <div class="page-wrap">
    <a-card class="main-card">
      <a-spin :spinning="processing" tip="解析与校验中...">
        <a-space direction="vertical" :size="24" style="width: 100%">
          <!-- 标题 -->
          <a-typography-title :level="4" class="page-title">
            价格合计与明细小计对比
          </a-typography-title>

          <!-- 配置区：列名称设置（第一步） -->
          <a-row :gutter="[16, 16]">
            <a-col :span="24">
              <a-typography-text strong>第一步：列名称配置</a-typography-text>
            </a-col>
            <a-col :xs="24" :sm="8">
              <div class="config-item">
                <label class="config-label">标记黄色列名称</label>
                <a-input
                  v-model:value="highlightColumnName"
                  placeholder="序号"
                  :disabled="processing"
                />
              </div>
            </a-col>
            <a-col :xs="24" :sm="8">
              <div class="config-item">
                <label class="config-label">合计价格列名称</label>
                <a-input
                  v-model:value="totalPriceColumnName"
                  placeholder="合计价格"
                  :disabled="processing"
                />
              </div>
            </a-col>
            <a-col :xs="24" :sm="8">
              <div class="config-item">
                <label class="config-label">明细价格列名称</label>
                <a-input
                  v-model:value="detailPriceColumnName"
                  placeholder="价格"
                  :disabled="processing"
                />
              </div>
            </a-col>
          </a-row>

          <a-row>
            <a-col :span="24">
              <div class="config-note">
                <p>标记黄色列名称：合计和明细之和不一致时会标记该列为黄。</p>
                <p>合计价格列名称：填写 Excel 里用于统计价格的列名称，作为分组字段。</p>
                <p>明细价格列名称：填写单行价格的列名称，会用于求和。</p>
              </div>
            </a-col>
          </a-row>

          <!-- 上传：点击或拖入，仅支持 Excel（第二步） -->
          <a-row>
            <a-col :span="24">
              <a-typography-text strong>第二步：上传 Excel 文件</a-typography-text>
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
            <p class="ant-upload-hint">仅支持上传 Excel 文件（.xlsx / .xls）</p>
            <p v-if="selectedFile" class="ant-upload-file-name">{{ selectedFile.name }}</p>
          </a-upload-dragger>

          <!-- 操作区：开始校验并下载 -->
          <div class="action-area">
            <a-button
              type="primary"
              size="large"
              :disabled="!selectedFile || processing"
              :loading="processing"
              @click="onStartValidate"
            >
              {{ processing ? '校验中...' : '开始校验并下载' }}
            </a-button>
          </div>

          <!-- 结果提示 -->
          <a-alert
            v-if="statusMessage"
            :message="statusMessage"
            :type="alertType"
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

interface ErrorGroup {
  groupKey: string;
  rows: number[];   // 该组涉及到的所有 Excel 行号
  sum: number;      // 价格之和
  total: number;    // 合计价格
}

const selectedFile = ref<File | null>(null);
const processing = ref(false);
const statusMessage = ref('');
const errors = ref<ErrorGroup[]>([]);

// 列名称配置
const highlightColumnName = ref('序号');
const totalPriceColumnName = ref('合计价格');
const detailPriceColumnName = ref('价格');

const alertType = computed(() => {
  if (!statusMessage.value) return 'info';
  if (statusMessage.value.startsWith('处理失败')) return 'error';
  if (statusMessage.value.includes('已下载')) return 'warning';
  return 'success';
});

function beforeUpload() {
  return false; // 阻止自动上传，仅本地处理
}

function onUploadChange(info: any) {
  const file = info?.file?.originFileObj ?? info?.file ?? null;
  selectedFile.value = file;
  statusMessage.value = file ? `已选择文件：${file.name}` : '';
  errors.value = [];
}

// 点击“开始校验并下载”
async function onStartValidate() {
  if (!selectedFile.value) return;

  processing.value = true;
  statusMessage.value = '正在解析与校验，请稍候...';
  errors.value = [];

  try {
    await validateAndDownload(selectedFile.value);
  } catch (err: any) {
    console.error(err);
    statusMessage.value = `处理失败：${err?.message || '未知错误'}`;
  } finally {
    processing.value = false;
  }
}

// 将 ExcelJS 的单元格值安全转为数字
function toNumber(value: any): number | null {
  if (value == null) return null;
  if (typeof value === 'number') return value;
  if (typeof value === 'string') {
    const v = value.trim();
    if (!v) return null;
    const n = Number(v);
    return Number.isNaN(n) ? null : n;
  }
  // 其他类型（公式、日期等）视为非数字，这里可以根据业务扩展
  return null;
}

// 根据行区间生成一个分组标签：优先用 A 列的序号，否则用行号区间
function getGroupLabel(ws: any, startRow: number, endRow: number): string {
  const seqCell = ws.getCell(startRow, 1);
  const raw = seqCell?.value;
  if (raw != null && String(raw).trim() !== '') {
    return String(raw).trim();
  }
  return startRow === endRow ? `行${startRow}` : `行${startRow}-${endRow}`;
}

// 主逻辑：解析、校验、标黄并下载
async function validateAndDownload(file: File) {
  const arrayBuffer = await file.arrayBuffer();

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(arrayBuffer);

  // 默认取第一个工作表，如需按名称可用 workbook.getWorksheet('Sheet1')
  const worksheet = workbook.worksheets[0];
  if (!worksheet) {
    throw new Error('Excel 中未找到任何工作表');
  }

  // 假设第一行是表头，从第二行开始为数据
  const DATA_START_ROW = 2;
  const HEADER_ROW = 1;

  // 根据列名称获取列号
  let highlightCol = -1;
  let totalPriceCol = -1;
  let detailPriceCol = -1;

  // 遍历表头行，找到对应的列号
  const headerRow = worksheet.getRow(HEADER_ROW);
  for (let c = 1; c <= (headerRow.cellCount || worksheet.columnCount); c++) {
    const cellValue = String(headerRow.getCell(c).value || '').trim();
    if (cellValue === highlightColumnName.value) {
      highlightCol = c;
    }
    if (cellValue === totalPriceColumnName.value) {
      totalPriceCol = c;
    }
    if (cellValue === detailPriceColumnName.value) {
      detailPriceCol = c;
    }
  }

  // 验证所有列是否都找到了
  const missingColumns: string[] = [];
  if (highlightCol === -1) {
    missingColumns.push(`标记黄色列"${highlightColumnName.value}"`);
  }
  if (totalPriceCol === -1) {
    missingColumns.push(`合计价格列"${totalPriceColumnName.value}"`);
  }
  if (detailPriceCol === -1) {
    missingColumns.push(`明细价格列"${detailPriceColumnName.value}"`);
  }

  // 如果有列找不到，抛出错误
  if (missingColumns.length > 0) {
    throw new Error(`Excel 表头中找不到以下列：${missingColumns.join('、')}。请检查列名称配置是否正确。`);
  }

  /**
   * 分组思路（按“合并价格”列来分组）：
   * - 查看每一行 D 列的单元格：
   *   - 如果 D 列是**合并单元格的主单元格**（cell.isMerged && cell.master.row === 当前行）：
   *     - 从主单元格开始向下扫描，同一个 master 的所有行构成一组；
   *   - 如果 D 列不是合并单元格（cell.isMerged === false）：
   *     - 这一行单独构成一组；
   * - 每一组内，对所有行的 C 列求和，与该组的 D 列“合并价格”比较。
   */

  // 为避免「上一次校验已涂黄，这次已修正仍然是黄的」，
  // 先清理当前工作表中数据区整行的黄色背景
  const lastRowNumber = worksheet.lastRow?.number ?? worksheet.rowCount;
  for (let r = DATA_START_ROW; r <= lastRowNumber; r++) {
    const row = worksheet.getRow(r);
    const maxCol = row.cellCount || worksheet.columnCount;
    for (let c = 1; c <= maxCol; c++) {
      const cell = row.getCell(c) as any;
      const fill = cell.fill;
      if (
        fill &&
        fill.type === 'pattern' &&
        fill.fgColor &&
        (fill.fgColor.argb === 'FFFFFF00' || fill.fgColor.argb === 'FFFF00')
      ) {
        cell.fill = undefined;
      }
    }
  }

  const errorGroups: ErrorGroup[] = [];

  // 使用 worksheet._merges 获取精确的合并区域，避免依赖 cell.master（可能未设置或为字符串）
  const merges = (worksheet as any)._merges as Record<string, { top: number; left: number; bottom: number; right: number }> | undefined;
  const usedRows = new Set<number>();

  if (merges && typeof merges === 'object') {
    Object.values(merges).forEach((range) => {
      if (range.left > totalPriceCol || range.right < totalPriceCol) return; // 不包含该列则跳过

      const startRow = Math.max(range.top, DATA_START_ROW);
      const endRow = range.bottom;
      if (endRow < DATA_START_ROW) return;

      const totalVal = toNumber(worksheet.getCell(range.top, totalPriceCol).value);
      if (totalVal == null) return;

      const rows: number[] = [];
      let sum = 0;
      for (let rowIdx = startRow; rowIdx <= endRow; rowIdx++) {
        rows.push(rowIdx);
        usedRows.add(rowIdx);
        const priceVal = toNumber(worksheet.getCell(rowIdx, detailPriceCol).value);
        if (priceVal != null) sum += priceVal;
      }

      const diffCents = Math.abs(
        Math.round(sum * 100) - Math.round(totalVal * 100),
      );
      if (diffCents >= 1) {
        errorGroups.push({
          groupKey: getGroupLabel(worksheet, startRow, endRow),
          rows,
          sum: Math.round(sum * 100) / 100,
          total: Math.round(totalVal * 100) / 100,
        });
      }
    });
  }

  // 未参与合并、但该列有值的单行
  for (let r = DATA_START_ROW; r <= lastRowNumber; r++) {
    if (usedRows.has(r)) continue;

    const dCell = worksheet.getCell(r, totalPriceCol);
    if (dCell.isMerged) continue; // 已在上面的合并里处理过（保险起见再跳过）

    const totalVal = toNumber(dCell.value);
    if (totalVal == null) continue;

    const priceVal = toNumber(worksheet.getCell(r, detailPriceCol).value);
    const sum = priceVal ?? 0;
    const diffCents = Math.abs(
      Math.round(sum * 100) - Math.round(totalVal * 100),
    );
    if (diffCents >= 1) {
      errorGroups.push({
        groupKey: getGroupLabel(worksheet, r, r),
        rows: [r],
        sum: Math.round(sum * 100) / 100,
        total: Math.round(totalVal * 100) / 100,
      });
    }
  }

  // 根据记录的行号进行标黄：汇总所有错误分组中的行号，根据配置的列标黄
  const rowsToHighlight = new Set<number>();
  errorGroups.forEach((g) => {
    g.rows.forEach((rowNum) => rowsToHighlight.add(rowNum));
  });
  console.log(rowsToHighlight);
  rowsToHighlight.forEach((rowNumber) => {
    const cell = worksheet.getCell(rowNumber, highlightCol);
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFFFFF00' }, // #FFFF00
    };
  });

  errors.value = errorGroups;

  if (errorGroups.length === 0) {
    statusMessage.value = '所有数据经过校验后都一致。';
  } else {
    statusMessage.value = '校验完成，已下载标注后的文件。';
  }

  // 仅在不一致时下载 Excel 文件
  if (errorGroups.length > 0) {
    const outBuffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([outBuffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    const originalName = file.name.replace(/\.xlsx?$/i, '');
    const outName = `${originalName}_校验后.xlsx`;
    saveAs(blob, outName);
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
.config-item {
  display: flex;
  flex-direction: column;
  gap: 8px;
}
.config-label {
  font-size: 12px;
  color: rgba(0, 0, 0, 0.65);
  font-weight: 500;
}
.config-note {
  background: #fffbe6; /* 淡黄色背景 */
  padding: 10px 12px;
  border-radius: 6px;
  border: 1px solid rgba(255, 200, 0, 0.12);
  font-size: 12px;
  color: rgba(0, 0, 0, 0.75);
  line-height: 1.6;
  margin-top: 6px;
}
.config-note p {
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

