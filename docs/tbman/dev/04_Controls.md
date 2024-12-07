---
title: 控制
keywords: VB6/VBA, WebView2浏览器控件, 可做采集
desc: VB6/VBA可用的WebView2浏览器控件(可做采集)
---

感谢你的耐心！接下来是 **控制功能部分 (Controls)** 的文档内容。我将继续提供每个控制功能的功能说明、参数说明和示例代码。

---

### **控制功能部分 (Controls)**

#### **1. SetVirtualHostNameToFolderMapping**
- **功能说明**：为指定的主机名创建虚拟文件夹映射，将指定主机名指向本地文件夹。可以通过此方法在 WebView2 中访问本地文件。
- **参数**：
  - `hostName` (String)：虚拟映射的主机名。
  - `folderPath` (String)：要映射的本地文件夹路径。
  - `accessKind` (wv2HostResourceAccessKind)：指定访问类型，默认允许访问。
- **示例代码**：
  ```php
  WebView21.SetVirtualHostNameToFolderMapping("example.local", "C:\MyLocalFolder", wv2HostResourceAccessKind.wv2ResourceAllow)
  ```

---

#### **2. ClearVirtualHostNameToFolderMapping**
- **功能说明**：移除通过 `SetVirtualHostNameToFolderMapping` 方法设置的虚拟文件夹映射。
- **参数**：
  - `hostName` (String)：要移除映射的主机名。
- **示例代码**：
  ```php
  WebView21.ClearVirtualHostNameToFolderMapping("example.local")
  ```

---

#### **3. PrintToPdf**
- **功能说明**：将当前文档保存为 PDF 文件。支持设置页面方向、缩放、边距等选项。
- **参数**：
  - `outputPath` (String)：输出的 PDF 文件路径。
  - `Orientation` (wv2PrintOrientation)：页面的打印方向，`wv2PrintPortrait` 或 `wv2PrintLandscape`。
  - `ScaleFactor` (Variant)：缩放因子，指定打印时页面的缩放比例。
  - `PageWidth`、`PageHeight`、`MarginTop`、`MarginBottom`、`MarginLeft`、`MarginRight` (Variant)：页面大小和边距。
  - `ShouldPrintBackgrounds` (Boolean)：是否打印背景。
  - `ShouldPrintSelectionOnly` (Boolean)：是否仅打印选中的部分。
  - `ShouldPrintHeaderAndFooter` (Boolean)：是否打印页眉和页脚。
  - `HeaderTitle`、`FooterUri` (Variant)：页眉和页脚内容。
- **示例代码**：
  ```php
  WebView21.PrintToPdf("C:\Output\document.pdf", wv2PrintOrientation.wv2PrintPortrait, 1.0, 8.5, 11, 1, 1, 1, 1, False, False, True, "My Document", "http://www.example.com")
  ```

---

#### **4. Reload**
- **功能说明**：重新加载当前页面，相当于按下 F5 键。
- **参数**：无。
- **示例代码**：
  ```php
  WebView21.Reload()
  ```

---

#### **5. Suspend**
- **功能说明**：暂停 WebView2 的处理和渲染，适用于类似标签页的功能。
- **参数**：无。
- **示例代码**：
  ```php
  WebView21.Suspend()
  ```

---

#### **6. Resume**
- **功能说明**：恢复 WebView2 的处理和渲染。只有在调用过 `Suspend` 后才有效。
- **参数**：无。
- **示例代码**：
  ```php
  WebView21.Resume()
  ```

---

#### **7. OpenDefaultDownloadDialog**
- **功能说明**：打开 WebView2 的默认下载管理器对话框。只有在支持该功能的 WebView2 中可用。
- **参数**：无。
- **示例代码**：
  ```php
  WebView21.OpenDefaultDownloadDialog()
  ```

---

#### **8. CloseDefaultDownloadDialog**
- **功能说明**：关闭 WebView2 的默认下载管理器对话框。只有在支持该功能的 WebView2 中可用。
- **参数**：无。
- **示例代码**：
  ```php
  WebView21.CloseDefaultDownloadDialog()
  ```

---

#### **9. OpenTaskManagerWindow**
- **功能说明**：打开 WebView2 的内置任务管理器对话框，用于调试和监控 WebView2 的运行。
- **参数**：无。
- **示例代码**：
  ```php
  WebView21.OpenTaskManagerWindow()
  ```

---

#### **10. AreDefaultScriptDialogsEnabled**
- **功能说明**：控制是否启用默认的 JavaScript 脚本对话框（如 `alert()`）。如果禁用，开发者可以通过监听 `ScriptDialogOpening` 事件来自定义对话框。
- **参数**：
  - `Value` (Boolean)：设置为 `True` 启用默认对话框，`False` 禁用默认对话框。
- **示例代码**：
  ```php
  WebView21.AreDefaultScriptDialogsEnabled = False
  ```

---

#### **11. AreDevToolsEnabled**
- **功能说明**：控制是否允许通过右键菜单或快捷键打开开发者工具。设置为 `True` 时，用户可以通过右键或快捷键打开 DevTools。
- **参数**：
  - `Value` (Boolean)：设置为 `True` 启用 DevTools，`False` 禁用 DevTools。
- **示例代码**：
  ```php
  WebView21.AreDevToolsEnabled = True
  ```

---

### 下一步
如果这些内容没有问题，我将继续处理文档的 **功能控制部分 (Feature Controls)**，包括一些其他控制和配置项。