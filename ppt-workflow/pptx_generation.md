# PPTX Generation and Editing

## 当前环境的优先工具

### 新建或一般编辑

优先使用已安装的官方 `Presentations` 技能和 `@oai/artifact-tool`：

- 按官方实现规范创建原生、可编辑的文字、表格、图表和图形。
- 使用官方 finalizer 进行包完整性、布局、字体、原生表格和图表检查。
- 不使用整页图片替代可编辑内容。

### 复杂已有 PPT 的保真编辑

在 Windows 且 Microsoft PowerPoint 可用时，优先考虑 PowerPoint COM 对复制件做原位修改，尤其适合：

- 只改颜色、填充、边框和阴影。
- 必须保留动画、母版、图片裁剪、嵌入对象和现有布局。
- 需要读取或写入 PowerPoint 原生属性。

编辑后仍应使用官方 Presentations 验证器和逐页渲染检查。不要把 COM 编辑等同于视觉检查。

### LibreOffice

用于渲染、格式兼容检查或没有 PowerPoint 时的有限转换。不要把复杂保真编辑默认交给 LibreOffice，因为它可能改变动画、字体或对象行为。

### python-pptx

仅作为简单、独立、无复杂动画和母版要求的兼容性备选。当前官方 Presentations 工作流要求使用 `@oai/artifact-tool` 时，不使用 `python-pptx` 替代。不要用它重存需要高保真的用户现有文件。

## 场景选择

| 任务 | 优先方法 |
| --- | --- |
| 新建可编辑 PPTX | 官方 Presentations + Artifact Tool |
| 新建原生表格/图表 | Artifact Tool |
| 复杂现有文件的小范围视觉修改 | PowerPoint COM + 结构快照 + 官方验证 |
| 只读检查或预览 | 官方渲染工具 / PowerPoint 导出 |
| 无 PowerPoint 环境的兼容转换 | 捆绑 LibreOffice |

## 不变量与版本管理

- 修改前创建备份。
- 候选文件放在私有构建目录，最终文件放在独立输出目录。
- 最终文件使用新名称，除非用户明确要求覆盖。
- 颜色批量替换时递归处理 slide、group、shape、fill、line、text、table 和 chart，但跳过图片像素本身。
- 修改后比较文字、几何、图片裁剪、页数、尺寸和动画元数据。

## 本地能力

当前环境已具备官方 Presentations 技能、Artifact Tool、PowerPoint COM、LibreOffice/渲染器和 PPTX 包校验器。通常不需要再下载第三方仓库。
