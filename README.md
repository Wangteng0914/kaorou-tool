# 🍖 南京烤肉烧烤店铺数据处理工具

上传区县Excel → 自动清洗/分类/排序 → 高德POI核验 → 导出飞书表格

---

## 🚀 一键部署到云端（手机随时用）

### 第一步：注册GitHub账号

1. 打开 https://github.com
2. 点右上角「Sign up」
3. 输入邮箱、密码、用户名，按提示完成注册
4. 登录后继续下一步

### 第二步：创建代码仓库

1. 登录GitHub后，点右上角 **「+」** → 选 **「New repository」**
2. 填写：
   - **Repository name**：`kaorou-tool`（随便起，英文就行）
   - **Public** 选上（免费）
   - 不用勾选其他选项
3. 点绿色按钮 **「Create repository」**

### 第三步：上传文件

1. 在刚创建的仓库页面，点 **「uploading an existing file」**（在蓝色提示框里）
2. 把下面这 **3个文件** 拖进去（或点「choose your files」选择）：
   - `kaorou_tool.py`
   - `requirements.txt`
   - `.streamlit` 文件夹（整个拖进去）
3. 拉到页面底部，点绿色按钮 **「Commit changes」**

> ⚠️ `.streamlit` 文件夹可能是隐藏的，如果看不到，在文件选择窗口按 `Ctrl+H` 显示隐藏文件。如果还是不行，就只上传前两个文件，不影响使用。

### 第四步：部署到Streamlit Cloud

1. 打开 https://share.streamlit.io
2. 点 **「Sign in with GitHub」**，用刚才的GitHub账号登录
3. 授权Streamlit访问你的GitHub（点「Authorize streamlit」）
4. 登录后点 **「New app」**
5. 填写：
   - **Repository**：选你刚创建的 `kaorou-tool`
   - **Branch**：`main`（默认就是）
   - **Main file path**：`kaorou_tool.py`
6. 点 **「Deploy!」**
7. 等待1-3分钟（页面会显示安装日志，看到 `You can now view your Streamlit app in your browser` 就成功了）

### 第五步：手机使用

1. 部署成功后，页面顶部会显示你的网址，比如：
   ```
   https://kaorou-tool-xxx.streamlit.app
   ```
2. 在手机浏览器打开这个网址
3. **添加到桌面**：
   - 苹果手机：Safari打开 → 底部分享按钮 →「添加到主屏幕」
   - 安卓手机：Chrome打开 → 右上角菜单 →「添加到主屏幕」或「安装应用」
4. 以后点桌面图标直接打开，和APP一样用

---

## 📱 使用方法

1. 打开工具后，左侧点 **「Browse files」** 上传区县Excel文件（可多选）
2. （可选）在 **「高德Web-Service API Key」** 输入你的高德Key
3. 选择 **「核验范围」**：仅烤肉店 / 全部店铺 / 不核验
4. 工具自动处理，每一步显示进度
5. 完成后点 **「📥 下载Excel文件」** 保存结果
6. 打开飞书 → 新建表格 → 导入 → 选择下载的Excel

---

## 🔑 高德API Key申请

1. 打开 https://console.amap.com/
2. 注册/登录账号
3. 左侧「应用管理」→「我的应用」→「创建新应用」
4. 应用名称随便填，类型选「工具」
5. 创建后点「添加Key」：
   - Key名称：随便填
   - 服务平台：**选「Web服务」**（很重要，选别的会报错）
6. 复制生成的Key，粘贴到工具里即可

> 免费额度：每日30万次调用，足够处理数千家店铺。

---

## 💻 本地运行（不想部署云端）

如果想在自己电脑上运行：

```bash
# 安装依赖
pip install streamlit pandas openpyxl pypinyin requests xlrd

# 运行
streamlit run kaorou_tool.py
```

浏览器自动打开 http://localhost:8501

---

## ❓ 常见问题

**Q: 部署失败怎么办？**
A: 看部署页面的错误日志。最常见的原因是 `requirements.txt` 没上传，或者文件名写错了（必须是 `kaorou_tool.py`）。

**Q: 手机打开网址显示「Please wait...」很久？**
A: 第一次打开Streamlit会冷启动，等30秒-1分钟就好。之后打开就快了。

**Q: 高德核验报错「USERKEY_PLAT_NOMATCH」？**
A: Key的服务平台选错了，必须选「Web服务」，重新创建一个Key。

**Q: 上传Excel后报错？**
A: 检查Excel是否能正常打开，第一行是否是表头（包含「店铺名称」「联系电话」「联系地址」等字段）。

**Q: 想更新工具代码怎么办？**
A: 把新的 `kaorou_tool.py` 重新上传到GitHub仓库覆盖旧文件，Streamlit会自动重新部署（等1-2分钟）。

---

## 📊 功能说明

- **数据清洗**：合并多文件 → 过滤仅11位手机号 → 按「店名+电话」去重
- **自动分类**：烤肉店（含"烤肉"）/ 烧烤店（含"烧烤"不含"烤肉"）/ 其他店
- **智能排序**：数字开头优先 → 拼音A-Z → 同名品牌分店连续排列
- **高德核验**：POI精准匹配，判定正常营业/暂停营业/已闭店
- **导出Excel**：带表头配色、冻结首行、大众点评超链接、手机号文本格式，直接导入飞书
