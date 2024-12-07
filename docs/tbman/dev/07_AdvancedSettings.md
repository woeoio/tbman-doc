---
title: 高级设置
keywords: VB6/VBA, WebView2浏览器控件, 可做采集
desc: VB6/VBA可用的WebView2浏览器控件(可做采集)
---

感谢你的耐心！现在是 **控制功能与高级设置部分 (Control Features & Advanced Settings)** 的文档内容。

---

### **控制功能与高级设置部分 (Control Features & Advanced Settings)**

#### **1. AreDefaultContextMenusEnabled**
- **功能说明**：控制是否启用默认的上下文菜单（右键菜单）。如果禁用，可以通过 `UserContextMenu` 事件来自定义右键菜单。
- **参数**：
  - `Value` (Boolean)：设置为 `True` 启用默认上下文菜单，`False` 禁用默认上下文菜单。
- **示例代码**：
  ```php
  ' 禁用默认上下文菜单
  WebView21.AreDefaultContextMenusEnabled = False
  
  ' 启用默认上下文菜单
  WebView21.AreDefaultContextMenusEnabled = True
  ```

---

#### **2. AreBrowserAcceleratorKeysEnabled**
- **功能说明**：控制是否启用浏览器加速键（如 F5 刷新、Ctrl+T 新标签页）。启用时，用户可以使用这些快捷键进行常见操作。
- **参数**：
  - `Value` (Boolean)：设置为 `True` 启用浏览器加速键，`False` 禁用加速键。
- **示例代码**：
  ```php
  ' 禁用浏览器加速键
  WebView21.AreBrowserAcceleratorKeysEnabled = False
  
  ' 启用浏览器加速键
  WebView21.AreBrowserAcceleratorKeysEnabled = True
  ```

---

#### **3. SupportsPdfFeatures**
- **功能说明**：返回一个布尔值，指示 WebView2 是否支持 PDF 功能。启用时，WebView2 可以直接展示 PDF 文件。
- **参数**：无。
- **示例代码**：
  ```php
  If WebView21.SupportsPdfFeatures Then
      MsgBox "支持 PDF 功能"
  Else
      MsgBox "不支持 PDF 功能"
  End If
  ```

---

#### **4. IsPinchZoomEnabled**
- **功能说明**：控制是否启用触摸设备上的捏合缩放功能。启用后，用户可以通过捏合手势进行页面缩放。
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

#### **5. SupportsAutoFillFeatures**
- **功能说明**：返回一个布尔值，指示 WebView2 是否支持自动填充功能。如果返回 `True`，表示 WebView2 支持自动填充（如自动填写表单中的用户名和密码）。
- **参数**：无。
- **示例代码**：
  ```php
  If WebView21.SupportsAutoFillFeatures Then
      MsgBox "支持自动填充功能"
  Else
      MsgBox "不支持自动填充功能"
  End If
  ```

---

#### **6. AreHostObjectsAllowed**
- **功能说明**：控制是否允许 WebView2 使用宿主对象（如 COM 对象）。如果设置为 `True`，JavaScript 可以访问和调用宿主对象的方法。
- **参数**：
  - `Value` (Boolean)：设置为 `True` 允许使用宿主对象，`False` 禁止使用宿主对象。
- **示例代码**：
  ```php
  ' 启用宿主对象
  WebView21.AreHostObjectsAllowed = True
  
  ' 禁用宿主对象
  WebView21.AreHostObjectsAllowed = False
  ```

---

#### **7. SupportsFolderMappingFeatures**
- **功能说明**：返回一个布尔值，指示 WebView2 是否支持文件夹映射功能。如果返回 `True`，则可以使用 `SetVirtualHostNameToFolderMapping` 设置主机名到文件夹的映射。
- **参数**：无。
- **示例代码**：
  ```php
  If WebView21.SupportsFolderMappingFeatures Then
      MsgBox "支持文件夹映射功能"
  Else
      MsgBox "不支持文件夹映射功能"
  End If
  ```

---

#### **8. IsStatusBarEnabled**
- **功能说明**：控制是否启用 WebView2 的状态栏。启用时，WebView2 会显示状态栏，用于显示页面状态或加载进度等信息。
- **参数**：
  - `Value` (Boolean)：设置为 `True` 启用状态栏，`False` 禁用状态栏。
- **示例代码**：
  ```php
  ' 启用状态栏
  WebView21.IsStatusBarEnabled = True
  
  ' 禁用状态栏
  WebView21.IsStatusBarEnabled = False
  ```

---

#### **9. SupportsTaskManagerFeatures**
- **功能说明**：返回一个布尔值，指示 WebView2 是否支持任务管理器功能。如果返回 `True`，则可以通过 `OpenTaskManagerWindow` 打开 WebView2 的内置任务管理器。
- **参数**：无。
- **示例代码**：
  ```php
  If WebView21.SupportsTaskManagerFeatures Then
      MsgBox "支持任务管理器功能"
  Else
      MsgBox "不支持任务管理器功能"
  End If
  ```

---

#### **10. SupportsAcceleratorKeysFeatures**
- **功能说明**：返回一个布尔值，指示 WebView2 是否支持浏览器加速键功能（如 F5 刷新、Ctrl+T 新标签页等）。
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

### 下一步
这部分是 WebView2 控制功能和高级设置的介绍。如果这部分没有问题，我将进行最后的文档整理并准备最终的输出。