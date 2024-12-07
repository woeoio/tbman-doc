---
title: 事件
keywords: VB6/VBA, WebView2浏览器控件, 可做采集
desc: VB6/VBA可用的WebView2浏览器控件(可做采集)
---

好的，我将分步为你提供文档内容，并确保每个部分都涵盖完整的解释。下面是第一部分——**事件部分**，我会逐个函数提供功能说明、参数说明和示例代码。

---

### **事件部分 (Events)**

#### **1. Create**
- **功能说明**：`Create` 事件在 WebView2 控件创建之前触发
- **参数**：无
- **示例代码**：
  ```php
  Private Sub WebView21_Create()
      MsgBox("WebView2 创建中...")
  End Sub
  ```

---

#### **2. Ready**
- **功能说明**：`Ready` 事件在 WebView2 控件成功创建并准备好使用时触发。
- **参数**：无
- **示例代码**：
  ```php
  Private Sub WebView21_Ready()
      MsgBox("WebView2 已准备就绪！")
  End Sub
  ```

---

#### **3. Error**
- **功能说明**：捕获 WebView2 控件初始化过程中发生的错误。
- **参数**：
  - `code` (Long)：错误代码。
  - `msg` (String)：错误信息。
- **示例代码**：
  ```php
  Private Sub WebView21_Error(ByVal code As Long, ByVal msg As String)
      MsgBox("错误代码: " & code & " 错误信息: " & msg)
  End Sub
  ```

---

#### **4. NavigationStarting**
- **功能说明**：`NavigationStarting` 事件在导航开始之前触发，允许开发者在导航前对请求进行修改或取消。
- **参数**：
  - `Uri` (String)：导航目标 URL。
  - `IsUserInitiated` (Boolean)：是否由用户触发导航。
  - `IsRedirected` (Boolean)：是否为重定向。
  - `RequestHeaders` (WebView2RequestHeaders)：请求头。
  - `Cancel` (ByRef Boolean)：设置为 `True` 可取消导航。
- **示例代码**：
  ```php
  Private Sub WebView21_NavigationStarting(ByVal Uri As String, ByVal IsUserInitiated As Boolean, ByVal IsRedirected As Boolean, ByVal RequestHeaders As Object, ByRef Cancel As Boolean)
      If Uri = "http://www.blockedwebsite.com" Then
          Cancel = True
          MsgBox("禁止访问该网站")
      End If
  End Sub
  ```

---

#### **5. NavigationComplete**
- **功能说明**：导航完成后触发该事件。
- **参数**：
  - `IsSuccess` (Boolean)：表示导航是否成功。
  - `WebErrorStatus` (Long)：导航失败时的错误状态码。
- **示例代码**：
  ```php
  Private Sub WebView21_NavigationComplete(ByVal IsSuccess As Boolean, ByVal WebErrorStatus As Long)
      If IsSuccess Then
          MsgBox("导航成功完成！")
      Else
          MsgBox("导航失败，错误码：" & WebErrorStatus)
      End If
  End Sub
  ```

---

#### **6. SourceChanged**
- **功能说明**：当 `DocumentURL` 属性更新时触发，通常用于页面导航完成后，更新 URL 栏等组件。
- **参数**：
  - `IsNewDocument` (Boolean)：是否为新文档。
- **示例代码**：
  ```php
  Private Sub WebView21_SourceChanged(ByVal IsNewDocument As Boolean)
      MsgBox("当前文档已更改：" & WebView21.DocumentURL)
  End Sub
  ```

---

### 下一步
如果这部分内容没有问题，我将继续为你提供方法和属性部分的说明。