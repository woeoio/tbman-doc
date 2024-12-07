---
title: 其他功能
keywords: VB6/VBA, WebView2浏览器控件, 可做采集
desc: VB6/VBA可用的WebView2浏览器控件(可做采集)
---

非常感谢你的耐心！我们已完成了大部分文档内容。接下来是 **总结与其他功能部分 (Summary & Other Features)** 的文档内容。

---

### **总结与其他功能部分 (Summary & Other Features)**

#### **1. UserContextMenu**
- **功能说明**：当用户右击 WebView2 控件时触发该事件。如果禁用了默认上下文菜单 (`AreDefaultContextMenusEnabled = False`)，则可以通过此事件自定义右键菜单。
- **参数**：
  - `X` (Single)：右键点击时的 X 坐标。
  - `Y` (Single)：右键点击时的 Y 坐标。
- **示例代码**：
  ```php
  Private Sub WebView21_UserContextMenu(X As Single, Y As Single)
      MsgBox "用户右键点击位置：X=" & X & ", Y=" & Y
  End Sub
  ```

---

#### **2. ScriptDialogOpening**
- **功能说明**：当 `AreDefaultScriptDialogsEnabled` 设置为 `False` 时，JavaScript 中的 `alert()`, `confirm()`, `prompt()` 等默认对话框会被拦截，并触发该事件。可以通过此事件自定义对话框。
- **参数**：
  - `ScriptDialogKind` (wv2ScriptDialogKind)：脚本对话框的类型（如 `alert`、`confirm` 等）。
  - `Accept` (ByRef Boolean)：设置为 `True` 可接受对话框，`False` 可拒绝对话框。
  - `ResultText` (String)：用于接收用户输入的文本。
  - `URI` (String)：触发对话框的 URL。
  - `Message` (String)：对话框中的消息。
  - `DefaultText` (String)：对话框中的默认文本。
- **示例代码**：
  ```php
  Private Sub WebView21_ScriptDialogOpening(ByVal ScriptDialogKind As wv2ScriptDialogKind, ByRef Accept As Boolean, ByVal ResultText As String, ByVal URI As String, ByVal Message As String, ByVal DefaultText As String)
      MsgBox "自定义脚本对话框: " & Message
      Accept = True ' 接受对话框
  End Sub
  ```

---

#### **3. SuspendCompleted**
- **功能说明**：在调用 `Suspend` 方法成功后触发的事件，表示 WebView2 已成功暂停其处理和渲染。
- **参数**：无。
- **示例代码**：
  ```php
  Private Sub WebView21_SuspendCompleted()
      MsgBox "WebView2 已成功暂停"
  End Sub
  ```

---

#### **4. SuspendFailed**
- **功能说明**：在调用 `Suspend` 方法失败时触发的事件。
- **参数**：无。
- **示例代码**：
  ```php
  Private Sub WebView21_SuspendFailed()
      MsgBox "WebView2 暂停失败"
  End Sub
  ```

---

#### **5. NewWindowRequested**
- **功能说明**：当 WebView2 尝试打开一个新窗口时触发此事件。通过设置 `IsHandled = True`，可以阻止 WebView2 默认行为，改为自定义窗口打开。
- **参数**：
  - `IsUserInitiated` (Boolean)：是否为用户触发。
  - `IsHandled` (ByRef Boolean)：如果设置为 `True`，则 WebView2 不会打开新窗口。
  - `Uri` (String)：请求的新 URL。
  - `HasPosition`、`HasSize`、`Left`、`Top`、`Width`、`Height` (Long)：新窗口的位置和大小。
  - `ShouldDisplayMenuBar`、`ShouldDisplayStatus`、`ShouldDisplayToolbar`、`ShouldDisplayScrollBars` (Long)：新窗口的功能设置。
- **示例代码**：
  ```php
  Private Sub WebView21_NewWindowRequested(ByVal IsUserInitiated As Boolean, ByRef IsHandled As Boolean, ByVal Uri As String, ByVal HasPosition As Long, ByVal HasSize As Long, ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long, ByVal ShouldDisplayMenuBar As Long, ByVal ShouldDisplayStatus As Long, ByVal ShouldDisplayToolbar As Long, ByVal ShouldDisplayScrollBars As Long)
      MsgBox "新窗口请求: " & Uri
      IsHandled = True ' 阻止默认打开新窗口
  End Sub
  ```

---

#### **6. DownloadStarting**
- **功能说明**：在用户点击下载链接时触发此事件。可以在此事件中修改下载路径或取消下载。
- **参数**：
  - `ResultFilePath` (ByRef String)：下载的目标文件路径。
  - `Cancel` (ByRef Boolean)：设置为 `True` 可取消下载。
  - `Handled` (ByRef Boolean)：设置为 `True` 处理下载，`False` 使用默认行为。
- **示例代码**：
  ```php
  Private Sub WebView21_DownloadStarting(ByRef ResultFilePath As String, ByRef Cancel As Boolean, ByRef Handled As Boolean)
      MsgBox "下载开始，目标文件路径: " & ResultFilePath
      ' 可修改文件路径或取消下载
      ResultFilePath = "C:\NewPath\DownloadedFile.exe"
      Cancel = False ' 允许下载继续
      Handled = True ' 处理了下载
  End Sub
  ```

---

### 最后整理与输出

所有内容已经整理完毕！文档涵盖了事件、方法、属性、控制功能、开发者工具、调试功能和其他特性等方面，已完整完成。

**接下来的步骤**：
- 我会将整个文档整理成一个文件，并提供给你下载。
- 如果你需要，我可以继续压缩并提供文件下载链接。

如果没有其他问题，我将立即准备文件并通知你。