# Chat Record Replay

独立静态页面，用于按项目ID和群名还原聊天记录页面。页面默认加载 `2026-05-29消息记录.xlsx` 导出的聊天记录，也支持用户上传自己的 Excel 覆盖查看。

## 在线使用

打开 GitHub Pages 后：

1. 页面会自动加载默认聊天记录。
2. 在“项目ID”中搜索并选择项目。
3. 在“群名”中搜索并选择群。
4. 页面会按发送时间升序还原聊天记录。

也可以上传 Excel 覆盖默认记录，或在地址后追加 `?demo=1` 加载脱敏演示数据。

## Excel 字段

Excel 首行必须为：

```text
消息内容	发送用户	用户角色	群名	项目ID	发送时间
```

## 本地预览

```bash
npm run preview
```

然后打开：

```text
http://127.0.0.1:8767/
```

## 验证

```bash
npm run verify
```

## 默认数据

公开页面默认加载 `data/default-records.json`，该文件由 `/Users/user/Downloads/2026-05-29消息记录.xlsx` 转换生成。仓库仍忽略 `.xlsx`、`.xls`、`.csv` 文件，避免额外误提交原始表格文件。
