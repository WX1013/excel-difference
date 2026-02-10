import { createApp } from 'vue';
import Antd from 'ant-design-vue';
import 'ant-design-vue/dist/reset.css';
import ExcelValidator from './ExcelValidator.vue';

const app = createApp(ExcelValidator);
app.use(Antd);
app.mount('#app');

