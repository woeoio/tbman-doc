---
title: 方法
keywords: VB6/VBA, WebView2浏览器控件, 可做采集
desc: VB6/VBA可用的WebView2浏览器控件(可做采集)
---

好的！接下来是 **方法部分 (Methods)** 的文档内容。我将继续提供每个方法的功能说明、参数说明和示例代码。

---

### **方法部分 (Methods)**

#### **1. Navigate**
- **功能说明**：导航到指定的 URL。
- **参数**：
  - `Url` (String)：要导航到的目标 URL。
- **示例代码**：
  ```php
  WebView21.Navigate("https://www.example.com")
  ```

---

#### **2. MoveFocus**
- **功能说明**：将焦点设置到 WebView 控件。
- **参数**：无。
- **示例代码**：
  ```php
  WebView21.MoveFocus()
  ```

---

#### **3. LoadHtml**
- **功能说明**：加载指定的 HTML 内容字符串。
- **参数**：
  - `htmlContent` (String)：要加载的 HTML 内容。
- **示例代码**：
  ```php
  WebView21.LoadHtml("<html><body><h1>Hello, World!</h1></body></html>")
  ```

---

#### **4. AddObject**
- **功能说明**：将 COM 对象暴露给 JavaScript 端，允许通过 `window.chrome.webview.hostObjects.ObjName` 访问。
- **参数**：
  - `ObjName` (String)：在 JavaScript 中访问对象的名称。
  - `Object` (Object)：要暴露给 JavaScript 的 COM 对象。
  - `UseDeferredInvoke` (Boolean)：是否使用延迟调用。
- **示例代码**：
  ```php
  WebView21.AddObject("myObject", myCustomObject, False)
  ```

---

#### **5. RemoveObject**
- **功能说明**：移除之前通过 `AddObject` 方法暴露的 COM 对象。
- **参数**：
  - `ObjName` (String)：要移除的对象名称。
- **示例代码**：
  ```php
  WebView21.RemoveObject("myObject")
  ```

---

#### **6. AddScriptToExecuteOnDocumentCreated**
- **功能说明**：在页面加载时注入 JavaScript 代码。仅对下一个页面导航生效。
- **参数**：
  - `jsCode` (String)：要执行的 JavaScript 代码。
- **示例代码**：
  ```php
  WebView21.AddScriptToExecuteOnDocumentCreated("console.log('Hello from injected script');")
  ```

---

#### **7. AddWebResourceRequestedFilter**
- **功能说明**：设置 URL 过滤器，当匹配的 URL 请求发生时，会触发 `WebResourceRequested` 事件。
- **参数**：
  - `sFilter` (String)：URL 过滤器，可以使用通配符 `*` 和 `?`。
  - `FilterContext` (wv2WebResourceContext)：请求上下文，指定过滤器适用的资源类型。
- **示例代码**：
  ```php
  WebView21.AddWebResourceRequestedFilter("https://www.example.com/*", WebView2RequestContext.wv2WebResourceContextScript)
  ```

---

#### **8. RemoveWebResourceRequestedFilter**
- **功能说明**：移除之前通过 `AddWebResourceRequestedFilter` 设置的 URL 过滤器。
- **参数**：
  - `sFilter` (String)：要移除的过滤器。
  - `FilterContext` (wv2WebResourceContext)：请求上下文。
- **示例代码**：
  ```php
  WebView21.RemoveWebResourceRequestedFilter("https://www.example.com/*", WebView2RequestContext.wv2WebResourceContextScript)
  ```

---

#### **9. OpenDevToolsWindow**
- **功能说明**：打开开发者工具窗口，允许开发者进行调试。
- **参数**：无。
- **示例代码**：
  ```php
  WebView21.OpenDevToolsWindow()
  ```

---

#### **10. CallDevToolsProtocolMethod**
- **功能说明**：向开发者工具发送协议消息，用于调试操作。
- **参数**：
  - `MethodName` (String)：要调用的开发者工具协议方法名称。
  - `ParamsAsJson` (String)：以 JSON 格式提供的方法参数。
  - `CustomEventId` (Variant)：可选，用户自定义的事件 ID，用于事件响应。
- **示例代码**：
  ```php
  WebView21.CallDevToolsProtocolMethod("Emulation.setScriptExecutionDisabled", "{""value"":true}")
  ```

---

#### **11. ExecuteScript**
- **功能说明**：执行 JavaScript 代码，但不等待结果返回，也不返回任何值。
- **参数**：
  - `jsCode` (String)：要执行的 JavaScript 代码。
- **示例代码**：
  ```php
  WebView21.ExecuteScript("alert('Hello from WebView!')")
  ```

---

#### **12. JsRun**
- **功能说明**：执行 JavaScript 函数，并等待结果返回。
- **参数**：
  - `FuncName` (String)：要执行的 JavaScript 函数名称。
  - `args()` (Variant)：可变参数，传递给 JavaScript 函数的参数。
- **示例代码**：
  ```php
  Dim result As Variant
  result = WebView21.JsRun("add", 1, 2)
  MsgBox result ' 输出 3
  ```

---

### 下一步
如果这一部分没有问题，我将继续处理文档中的 **属性部分 (Properties)**。