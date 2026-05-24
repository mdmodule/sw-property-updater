' ******************************************************************************
' SolidWorks 2025 属性智能更新宏 (标准版)
' 功能：
' 1. 保留原有"设计日期"，若无则设为今日。
' 2. 无论文件名是否规范，均全量清理并按 1-16 顺序重建属性表。
' 3. 文件名合规（名称+代号+版本）则自动拆分；不合规（无+号）则弹出警告并留空关键项。
' ******************************************************************************

Option Explicit

Sub main()
 Dim swApp As SldWorks.SldWorks
 Dim swModel As SldWorks.ModelDoc2
 Dim swCustPropMgr As SldWorks.CustomPropertyManager
 Dim fullPath As String
 Dim fileName As String
 Dim fileNameParts As Variant
 Dim vPropNames As Variant
 Dim i As Long
 Dim docType As Long

 ' 拆分变量
 Dim strName As String, strNumber As String, strRev As String
 Dim isFormatError As Boolean: isFormatError = False

 ' 日期处理变量
 Dim oldDate As String
 Dim valOut As String
 Dim resValOut As String

 Set swApp = Application.SldWorks
 Set swModel = swApp.ActiveDoc

 ' 检查是否有打开的文件
 If swModel Is Nothing Then
 MsgBox "请先打开一个零件或装配体文件！", vbCritical, "错误"
 Exit Sub
 End If

 ' 1. 获取基础信息
 docType = swModel.GetType
 fullPath = swModel.GetPathName

 If fullPath = "" Then
 fileName = swModel.GetTitle
 Else
 fileName = Mid(fullPath, InStrRev(fullPath, "\") + 1)
 fileName = Left(fileName, InStrRev(fileName, ".") - 1)
 End If

 ' 2. 文件名拆分逻辑 (以 + 号为界)
 fileNameParts = Split(fileName, "+")

 If UBound(fileNameParts) >= 2 Then
 ' 符合规则：名称+代号+版本
 strName = Trim(fileNameParts(0))
 strNumber = Trim(fileNameParts(1))
 strRev = Trim(fileNameParts(2))
 Else
 ' 不符合规则：触发警告，名称填入全名，其余留空
 isFormatError = True
 strName = fileName
 strNumber = ""
 strRev = ""
 End If

 ' 3. 获取属性管理器
 Set swCustPropMgr = swModel.Extension.CustomPropertyManager("")

 ' 4. 提取原有日期 (如果存在)
 swCustPropMgr.Get6 "设计日期", False, valOut, resValOut, False, False
 If Trim(resValOut) <> "" Then
 oldDate = Trim(resValOut)
 Else
 oldDate = CStr(Date)
 End If

 ' 5. 全量清理旧属性 (确保排序从第一行开始)
 vPropNames = swCustPropMgr.GetNames
 If Not IsEmpty(vPropNames) Then
 For i = 0 To UBound(vPropNames)
 swCustPropMgr.Delete vPropNames(i)
 Next i
 End If

 ' 6. 重新按 1-16 顺序精准写入
 ' 参数 1 表示 swCustomPropertyReplaceValue，即强制覆盖

 ' --- 基础信息 ---
 swCustPropMgr.Add3 "名称", 30, strName, 1
 swCustPropMgr.Add3 "零部件号", 30, strNumber, 1
 swCustPropMgr.Add3 "图样代号", 30, strNumber, 1

 ' --- 物理属性 ---
 If docType = 1 Then ' 零件
 swCustPropMgr.Add3 "材料", 30, """SW-Material""", 1
 Else ' 装配体
 swCustPropMgr.Add3 "材料", 30, "", 1
 End If
 swCustPropMgr.Add3 "质量", 30, """SW-Mass""", 1

 ' --- 状态与设计信息 ---
 swCustPropMgr.Add3 "单位", 30, "Pcs", 1
 swCustPropMgr.Add3 "设计", 30, "Abel", 1
 swCustPropMgr.Add3 "设计日期", 30, oldDate, 1

 ' --- 审核与流程占位 ---
 swCustPropMgr.Add3 "审核", 30, "", 1
 swCustPropMgr.Add3 "审核日期", 30, "", 1
 swCustPropMgr.Add3 "工艺", 30, "", 1
 swCustPropMgr.Add3 "工艺日期", 30, "", 1
 swCustPropMgr.Add3 "批准", 30, "", 1
 swCustPropMgr.Add3 "批准日期", 30, "", 1

 ' --- 版本与备注 ---
 swCustPropMgr.Add3 "版本号", 30, strRev, 1
 swCustPropMgr.Add3 "备注", 30, "", 1

 ' 7. 刷新文档
 swModel.ForceRebuild3 True

 ' 8. 最终反馈
 If isFormatError Then
 MsgBox "【警告】：文件名格式不符合规范（名称+代号+版本）！" & vbCrLf & _
 "属性表已按标准顺序强制初始化，请手动补全"代号"和"版本号"。", vbExclamation, "格式错误"
 Else
 MsgBox "属性更新成功！" & vbCrLf & "使用日期: " & oldDate, vbInformation, "执行完毕"
 End If
End Sub
