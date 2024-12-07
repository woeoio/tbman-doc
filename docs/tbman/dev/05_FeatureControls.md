---
title: 功能
keywords: VB6/VBA, WebView2浏览器控件, 可做采集
desc: VB6/VBA可用的WebView2浏览器控件(可做采集)
---

感谢你的耐心！接下来是 **功能控制部分 (Feature Controls)** 的文档内容。我将继续提供每个控制功能的功能说明、参数说明和示例代码。

---

### **功能控制部分 (Feature Controls)**

#### **1. IsScriptEnabled**
- **功能说明**：获取或设置是否启用 WebView2 中的 JavaScript。设置为 `False` 时，WebView2 将禁用 JavaScript 的执行。
- **参数**：
  - `Value` (Boolean)：设置为 `True` 启用 JavaScript，`False` 禁用 JavaScript。
- **示例代码**：
  ```php
  ' 禁用 JavaScript 执行
  WebView21.IsScriptEnabled = False
  
  ' 启用 JavaScript 执行
  WebView21.IsScriptEnabled = True
  ```

---

#### **2. IsWebMessageEnabled**
- **功能说明**：控制是否启用 `PostWebMessage` 功能，使 WebView2 可以与 JavaScript 端进行消息传递。设置为 `True` 启用消息传递。
- **参数**：
  - `Value` (Boolean)：设置为 `True` 启用消息传递，`False` 禁用消息传递。
- **示例代码**：
  ```php
  ' 启用 WebMessage 功能
  WebView21.IsWebMessageEnabled = True
  
  ' 禁用 WebMessage 功能
  WebView21.IsWebMessageEnabled = False
  ```

---

#### **3. AreHostObjectsAllowed**
- **功能说明**：控制是否允许在 WebView2 中暴露宿主对象。启用此功能后，JavaScript 可以通过 `chrome.webview.hostObjects` 访问宿主对象。
- **参数**：
  - `Value` (Boolean)：设置为 `True` 允许暴露宿主对象，`False` 禁止暴露宿主对象。
- **示例代码**：
  ```php
  ' 允许暴露宿主对象
  WebView21.AreHostObjectsAllowed = True
  
  ' 禁止暴露宿主对象
  WebView21.AreHostObjectsAllowed = False
  ```

---

#### **4. IsZoomControlEnabled**
- **功能说明**：控制是否允许用户通过快捷键（Ctrl + 鼠标滚轮、Ctrl + 加号、Ctrl + 减号）来调整页面缩放。设置为 `False` 时，禁用该功能。
- **参数**：
  - `Value` (Boolean)：设置为 `True` 启用缩放控制，`False` 禁用缩放控制。
- **示例代码**：
  ```php
  ' 启用缩放控制
  WebView21.IsZoomControlEnabled = True
  
  ' 禁用缩放控制
  WebView21.IsZoomControlEnabled = False
  ```

---

#### **5. IsBuiltInErrorPageEnabled**
- **功能说明**：控制是否启用 WebView2 内置的错误页面（如 404 页面）。如果禁用该功能，WebView2 会显示空白页面或开发者提供的错误页面。
- **参数**：
  - `Value` (Boolean)：设置为 `True` 启用内置错误页面，`False` 禁用内置错误页面。
- **示例代码**：
  ```php
  ' 启用内置错误页面
  WebView21.IsBuiltInErrorPageEnabled = True
  
  ' 禁用内置错误页面
  WebView21.IsBuiltInErrorPageEnabled = False
  ```

---

#### **6. SupportsAcceleratorKeysFeatures**
- **功能说明**：返回一个布尔值，指示 WebView2 是否支持加速键功能（如 F5 刷新）。如果返回 `True`，表示 WebView2 支持加速键。
- **参数**：无。
- **示例代码**：
  ```php
  If WebView21.SupportsAcceleratorKeysFeatures Then
      MsgBox "支持加速键功能"
  Else
      MsgBox "不支持加速键功能"
  End If
  ```

---

#### **7. IsPasswordAutoSaveEnabled**
- **功能说明**：控制是否启用自动保存密码功能。如果启用，WebView2 会自动保存和填充表单中的密码信息。
- **参数**：
  - `Value` (Boolean)：设置为 `True` 启用密码自动保存，`False` 禁用该功能。
- **示例代码**：
  ```php
  ' 启用密码自动保存
  WebView21.IsPasswordAutoSaveEnabled = True
  
  ' 禁用密码自动保存
  WebView21.IsPasswordAutoSaveEnabled = False
  ```

---

#### **8. IsGeneralAutoFillEnabled**
- **功能说明**：控制是否启用自动填充功能。如果启用，WebView2 会自动填充表单中的信息（如用户名、地址等）。
- **参数**：
  - `Value` (Boolean)：设置为 `True` 启用表单自动填充，`False` 禁用该功能。
- **示例代码**：
  ```php
  ' 启用表单自动填充
  WebView21.IsGeneralAutoFillEnabled = True
  
  ' 禁用表单自动填充
  WebView21.IsGeneralAutoFillEnabled = False
  ```

---

#### **9. IsPinchZoomEnabled**
- **功能说明**：控制是否启用触摸设备上的捏合缩放功能。启用后，用户可以使用捏合手势来放大或缩小页面。
- **参数**：
  - `Value` (Boolean)：设置为 `True` 启用捏合缩放，`False` 禁用捏合缩放。
- **示例代码**：
  ```php
  ' 启用捏合缩放
  WebView21.IsPinchZoomEnabled = True
  
  ' 禁用捏合缩放
  WebView21.IsPinchZoomEnabled = False
  ```

---

#### **10. IsSwipeNavigationEnabled**
- **功能说明**：控制是否启用触摸设备上的滑动导航功能。启用后，用户可以通过滑动手势进行页面前进或后退。
- **参数**：
  - `Value` (Boolean)：设置为 `True` 启用滑动导航，`False` 禁用滑动导航。
- **示例代码**：
  ```php
  ' 启用滑动导航
  WebView21.IsSwipeNavigationEnabled = True
  
  ' 禁用滑动导航
  WebView21.IsSwipeNavigationEnabled = False
  ```

---

### 下一步
如果这些功能控制部分没有问题，我会整理文档并准备最终的输出。如果你有其他问题或需要进一步的帮助，请告诉我！