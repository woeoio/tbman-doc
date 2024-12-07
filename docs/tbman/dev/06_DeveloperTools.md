---
title: 开发者调试
keywords: VB6/VBA, WebView2浏览器控件, 可做采集
desc: VB6/VBA可用的WebView2浏览器控件(可做采集)
---

非常感谢你的耐心！接下来是 **开发者工具与调试功能部分 (Developer Tools and Debugging Features)** 的文档内容。

---

### **开发者工具与调试功能部分 (Developer Tools and Debugging Features)**

#### **1. JsAsyncResult**
- **功能说明**：当 `JsRunAsync` 函数执行完成时，该事件会被触发。它提供了 JavaScript 执行结果的异步返回，可以通过 `Token` 来区分不同的调用。
- **参数**：
  - `Result` (Variant)：JavaScript 执行结果，通常是返回的值。
  - `Token` (LongLong)：调用时传递的唯一标识符，便于标识不同的异步请求。
  - `ErrString` (String)：错误信息（如果发生错误）。
- **示例代码**：
  ```php
  Private Sub WebView21_JsAsyncResult(ByVal Result As Variant, Token As LongLong, ErrString As String)
      If ErrString = "" Then
          MsgBox "JavaScript 执行结果：" & Result
      Else
          MsgBox "JavaScript 错误：" & ErrString
      End If
  End Sub
  ```

---

#### **2. JsMessage**
- **功能说明**：当 JavaScript 端通过 `window.chrome.webview.postMessage(...)` 发送消息时，`JsMessage` 事件会触发。此功能需要启用 `IsWebMessageEnabled`。
- **参数**：
  - `Message` (Variant)：来自 JavaScript 的消息内容，可以是任意类型的数据。
- **示例代码**：
  ```php
  Private Sub WebView21_JsMessage(ByVal Message As Variant)
      MsgBox "收到来自 JavaScript 的消息：" & Message
  End Sub
  ```

---

#### **3. DevToolsProtocolResponse**
- **功能说明**：当通过 `CallDevToolsProtocolMethod` 调用开发者工具协议方法时，`DevToolsProtocolResponse` 事件会接收到方法的响应。
- **参数**：
  - `CustomEventId` (Variant)：自定义的事件 ID，用于关联响应。
  - `JsonResponse` (String)：开发者工具协议方法的 JSON 响应。
- **示例代码**：
  ```php
  Private Sub WebView21_DevToolsProtocolResponse(ByVal CustomEventId As Variant, ByVal JsonResponse As String)
      MsgBox "DevTools 响应：" & JsonResponse
  End Sub
  ```

---

#### **4. OpenDevToolsWindow**
- **功能说明**：打开 WebView2 的开发者工具窗口。可以用于调试和查看页面内部内容。
- **参数**：无。
- **示例代码**：
  ```php
  WebView21.OpenDevToolsWindow()
  ```

---

#### **5. CallDevToolsProtocolMethod**
- **功能说明**：通过调用 WebView2 的开发者工具协议方法来与浏览器进行交互。例如，可以禁用脚本执行、模拟设备环境等。
- **参数**：
  - `MethodName` (String)：协议方法名称（如 `Emulation.setScriptExecutionDisabled`）。
  - `ParamsAsJson` (String)：以 JSON 格式传递的参数。
  - `CustomEventId` (Variant)：可选的自定义事件 ID，用于关联响应。
- **示例代码**：
  ```php
  WebView21.CallDevToolsProtocolMethod("Emulation.setScriptExecutionDisabled", "{""value"":true}")
  ```

---

#### **6. JsRunAsync**
- **功能说明**：异步执行 JavaScript 代码，并返回一个唯一的 `Token` 来标识该异步执行。执行完成后，`JsAsyncResult` 事件会触发返回结果。
- **参数**：
  - `FuncName` (String)：要执行的 JavaScript 函数名称。
  - `args()` (Variant)：可变参数，传递给 JavaScript 函数的参数。
- **示例代码**：
  ```php
  Dim Token As LongLong
  Token = WebView21.JsRunAsync("myFunction", "param1", "param2")
  MsgBox "异步调用的 Token：" & Token
  ```

---

#### **7. JsRun**
- **功能说明**：执行 JavaScript 函数并等待结果返回。适用于需要同步返回结果的情况。
- **参数**：
  - `FuncName` (String)：要执行的 JavaScript 函数名称。
  - `args()` (Variant)：可变参数，传递给 JavaScript 函数的参数。
- **示例代码**：
  ```php
  Dim result As Variant
  result = WebView21.JsRun("sum", 5, 10)
  MsgBox "执行结果：" & result ' 输出 15
  ```

---

### 下一步
这部分是 WebView2 的开发者工具和调试功能的介绍。如果这一部分没有问题，我将整理并提供最终的完整文档。