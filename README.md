# 滴滴发票行程单一键下载

## 项目目的
本项目旨在实现一键下载和归档滴滴发票。用户通过本插件，可以自动从 Outlook 中下载滴滴发票相关邮件的附件，进行 OCR 识别后重命名，并归档到指定的文件夹。

---

## 运作原理
本项目的运作分为以下几个步骤：

1. **邮件选择与附件提取**：
   - 通过 Outlook 的 VBE 提取用户选中的邮件。
   - 检查邮件标题是否包含关键字（如 `行程报销单`），提取附件并保存至临时文件夹。

2. **OCR 识别**：
   - 调用百度 OCR 接口对提取的发票文件进行识别。
   - 自动提取发票中的关键信息（至少需要 `通用文字识别（高精度版）` API）。

3. **文件重命名与归档**：
   - 使用正则表达式提取的关键信息，重命名附件。
   - 将重命名后的文件转移到用户在配置文件中指定的目标文件夹。

4. **配置管理**：
   - 提供配置窗体，允许用户设置目标文件夹以及百度 OCR API 的 `clientId` 和 `clientSecret`。

---

## 安装方法

### 方法 1：使用 VBS 自动安装

1. 确保已安装 Microsoft Outlook。
2. 双击运行 `Setup.vbs` 脚本。
3. 安装完成后，重新启动 Outlook，您将在工具栏中看到相关插件菜单。

### 方法 2：手动导入到 VBE

1. 打开 Microsoft Outlook。
2. 按下 `Alt + F11` 打开 VBA 编辑器。
3. 在菜单中选择 `File > Import File`，依次导入以下文件：
   - `DiDi_invoice.bas`
   - `SubFunction.bas`
   - `ConfigForm.frm` 和 `ConfigForm.frx`
   - `JsonConverter.bas`
4. 确保所有模块都正确加载后，保存项目。
5. 回到 Outlook 主界面，在工具栏中添加菜单以调用 `ShowConfigForm` 方法。

### 创建菜单

1. 在 Outlook 的 Ribbon 菜单栏上右键选择 `Customize the Ribbon`。
2. 点击 `New Tab`，为其命名，如 `滴滴发票管理`。
3. 在新建的 Tab 下创建一个 Group，命名为 `操作`。
4. 将对应的 VBA 方法（如 `ProcessMailAndCallOCR` 和 `ShowConfigForm`）分配为按钮。

---

## 致谢
本项目部分功能基于以下开源项目实现：
- [VBA-JSON](https://github.com/VBA-tools/VBA-JSON) 用于 JSON 数据解析。

感谢开源社区提供的工具与灵感！

---

## 许可协议
本项目采用 MIT 协议：
- 您可以自由复制、修改和分发，但请在使用时保留原作者信息。
- 引用或调用时请注明出处。

