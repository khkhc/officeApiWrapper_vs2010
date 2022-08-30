
#include "Cword.h"


CmyWord::CmyWord(){
	InitCOM();
}

CmyWord::~CmyWord(){
	//释放资源最好从 小到大的顺序来释放。这个和c里面一些释放资源的道理是一样的

	//和c+= 先析构儿子 再析构父亲是一样的。
	CoUninitialize();
	font.ReleaseDispatch();
	range.ReleaseDispatch();
	tab.ReleaseDispatch();
	doc.ReleaseDispatch();
	docs.ReleaseDispatch();
	app.ReleaseDispatch();
	sel.ReleaseDispatch();
}

BOOL CmyWord::InitCOM(){
	if (!CoInitialize(NULL)){
		AfxMessageBox(_T("初始化com库失败"));
		return 0;
	}
	else{
		return TRUE;
	}
}
BOOL CmyWord::CreateAPP(){
	if (!app.CreateDispatch(_T("Word.Application"))){
		AfxMessageBox(_T("你没有安装OFFICE"));
		return FALSE;
	}
	else{
		app.SetVisible(TRUE);
		return TRUE;
	}
}
//我的类默认是打开的，而Word 中默认看不见的。

void CmyWord::ShowApp(BOOL flag){
	if (!app.m_lpDispatch){
		AfxMessageBox(_T("你还没有获得Word对象"));
		return;
	}
	else{
		app.SetVisible(flag);
	}
}
BOOL CmyWord::CreateDocument(){
	if (!app.m_lpDispatch){
		AfxMessageBox(_T("Application为空,Documents创建失败!"), MB_OK | MB_ICONWARNING);
		return FALSE;
	}
	else{
		docs = app.GetDocuments();
		if (docs.m_lpDispatch == NULL){
			AfxMessageBox(_T("创建DOCUMENTS 失败"));
			return FALSE;
		}
		else{
			CComVariant Template(_T(""));//创建一个空的模版
			CComVariant NewTemplate(false);
			CComVariant DocumentType(0);
			CComVariant Visible;//不处理 用默认值
			doc = docs.Add(&Template, &NewTemplate, &DocumentType, &Visible);
			if (!doc.m_lpDispatch){
				AfxMessageBox(_T("创建word失败"));
				return FALSE;
			}
			else{
				sel = app.GetSelection();//获得当前Word操作。开始认为是在doc获得selection。仔细想一下确实应该是Word的接口点
				if (!sel.m_lpDispatch){
					AfxMessageBox(_T("selection 获取失败"));
					return FALSE;
				}
				else{
					return TRUE;
				}
			}
		}
	}
}

BOOL CmyWord::Create(){
	if (CreateAPP()){
		if (CreateDocument()){
			return TRUE;
		}
		else
			return FALSE;
	}
	else
		return FALSE;
}

BOOL CmyWord::Open(CString FileName, BOOL ReadOnly /* = FALSE */, BOOL AddToRecentFiles /* = FALSE */){
	CComVariant Read(ReadOnly);
	CComVariant AddToR(AddToRecentFiles);
	CComVariant Name(FileName);
	COleVariant vTrue((short)TRUE), vFalse((short)FALSE);
	COleVariant varstrNull(_T(""));
	COleVariant varZero((short)0);
	COleVariant varTrue(short(1), VT_BOOL);
	COleVariant varFalse(short(0), VT_BOOL);
	COleVariant vOpt((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	if (!app.m_lpDispatch){
		if (CreateAPP() == FALSE){
			return FALSE;
		}
	}
	if (!docs.m_lpDispatch){
		docs = app.GetDocuments();
		if (!docs.m_lpDispatch){

			AfxMessageBox(_T("DocuMent 对象创建失败"));
			return FALSE;
		}
	}
	CComVariant format(0);//打开方式 0 为doc的打开方式
	doc = docs.Open(&Name, varFalse, &Read, &AddToR, vOpt, vOpt,
		vFalse, vOpt, vOpt, &format, vOpt, vTrue, vOpt, vOpt, vOpt, vOpt);
	if (!doc.m_lpDispatch){
		AfxMessageBox(_T("文件打开失败"));
		return FALSE;
	}
	else{
		sel = app.GetSelection();
		if (!sel.m_lpDispatch){
			AfxMessageBox(_T("打开失败"));
			return FALSE;
		}
		return TRUE;
	}
}
BOOL CmyWord::Save(){
	if (!doc.m_lpDispatch){
		AfxMessageBox(_T("Documents 对象都没有建立 保存失败"));
		return FALSE;
	}
	else{
		doc.Save();
		return TRUE;
	}
}
BOOL CmyWord::SaveAs(CString FileName, int SaveType/* =0 */){
	CComVariant vTrue(TRUE);
	CComVariant vFalse(FALSE);
	CComVariant vOpt;
	CComVariant cFileName(FileName);
	CComVariant FileFormat(SaveType);
	doc = app.GetActiveDocument();
	if (!doc.m_lpDispatch){
		AfxMessageBox(_T("Document 对象没有建立 另存为失败"));
		return FALSE;
	}
	else{
		//最好按照宏来写 不然可能出现问题、 毕竟这个是微软写的
		/*ActiveDocument.SaveAs FileName:="xiaoyuer.doc", FileFormat:= _

	wdFormatDocument, LockComments:=False, Password:="", AddToRecentFiles:= _

	True, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts:= _

	False, SaveNativePictureFormat:=False, SaveFormsData:=False, _

		SaveAsAOCELetter:=False*/
		doc.SaveAs(&cFileName, &FileFormat, &vFalse, COleVariant(_T("")), &vTrue,
			COleVariant(_T("")), &vFalse, &vFalse, &vFalse, &vFalse, &vFalse, &vOpt, &vOpt, &vOpt, &vOpt, &vOpt);
	}
	return TRUE;
}
BOOL CmyWord::Close(BOOL SaveChange/* =FALSE */){
	CComVariant vTrue(TRUE);
	CComVariant vFalse(FALSE);
	CComVariant vOpt;
	CComVariant cSavechage(SaveChange);
	if (!doc.m_lpDispatch){
		AfxMessageBox(_T("_Document 对象获取失败,关闭操作失败"));
		return FALSE;
	}
	else{
		if (TRUE == SaveChange){
			Save();
		}
		//下面第一个参数填vTrue 会出现错误，可能是后面的参数也要对应的变化
		//但vba 没有给对应参数 我就用这种方法来保存
		doc.Close(&vFalse, &vOpt, &vOpt);
	}
	return TRUE;
}
void CmyWord::WriteText(CString Text){
	sel.TypeText(Text);
}
void CmyWord::NewLine(int nCount/* =1 */){
	if (nCount <= 0){
		nCount = 0;
	}
	else{
		for (int i = 0; i < nCount; i++){
			sel.TypeParagraph();//新建一段
		}
	}
}
void CmyWord::WriteTextNewLineText(CString Text, int nCount/* =1 */){
	NewLine(nCount);
	WriteText(Text);
}
void CmyWord::SetFont(BOOL Blod, BOOL Italic/* =FALSE */, BOOL UnderLine/* =FALSE */){
	if (!sel.m_lpDispatch){
		AfxMessageBox(_T("编辑对象失败,导致字体不能设置"));
		return;
	}
	else{
		sel.SetText(_T("F"));
		font = sel.GetFont();//获得字体编辑对象;
		font.SetBold(Blod);
		font.SetItalic(Italic);
		font.SetUnderline(UnderLine);
		sel.SetFont(font);
	}
}
void CmyWord::SetFont(CString FontName, int FontSize/* =9 */, long FontColor/* =0 */, long FontBackColor/* =0 */){
	if (!sel.m_lpDispatch){
		AfxMessageBox(_T("Select 为空,字体设置失败!"));
		return;
	}
	//这里只是为了获得一个对象，因为没有对象你哪里来的设置呢.
	//因为是用GetFont来获取的对象的。
	//所以用SetText来获得字体属性
	sel.SetText(_T("a"));
	font = sel.GetFont();//获取字体对象
	font.SetSize(20);
	font.SetName(FontName);
	font.SetColor(FontColor);
	sel.SetFont(font);//选择对象
}
void CmyWord::SetTableFont(int Row, int Column, CString FontName, int FontSize/* =9 */, long FontColor/* =0 */, long FontBackColor/* =0 */){
	Cell c = tab.Cell(Row, Column);
	c.Select();
	_Font ft = sel.GetFont();
	ft.SetName(FontName);
	ft.SetSize(FontSize);
	ft.SetColor(FontColor);
	Range r = sel.GetRange();
	r.SetHighlightColorIndex(FontBackColor);
}
void CmyWord::CreateTable(int Row, int Column)
{
	doc = app.GetActiveDocument();
	Tables tbs = doc.GetTables();
	CComVariant Vtrue(short(TRUE)), Vfalse(short(FALSE));
	if (!tbs.m_lpDispatch){
		AfxMessageBox(_T("创建表格对象失败"));
		return;
	}
	else{
		tbs.Add(sel.GetRange(), Row, Column, &Vtrue, &Vfalse);
		tab = tbs.Item(1);//如果有多个表格可以通过这个来找到表格对象。
	}
}
void CmyWord::WriteCellText(int Row, int Column, CString Text){
	Cell c = tab.Cell(Row, Column);
	c.Select();//选择表格中的单元格
	sel.TypeText(Text);
}
void CmyWord::SetParaphformat(int Alignment){
	_ParagraphFormat p = sel.GetParagraphFormat();
	p.SetAlignment(Alignment);
	sel.SetParagraphFormat(p);
}
void CmyWord::FindWord(CString FindW, CString RelWord){
	sel = app.GetSelection();
	Find myFind = sel.GetFind();
	if (!myFind.m_lpDispatch){
		AfxMessageBox(_T("获取Find 对象失败"));
		return;
	}
	else{
		//下面三行是按照vba 写的
		myFind.ClearFormatting();
		Replacement repla = myFind.GetReplacement();
		repla.ClearFormatting();
		COleVariant Text(FindW);
		COleVariant re(RelWord);
		COleVariant vTrue((short)TRUE), vFalse((short)FALSE);
		COleVariant vOpt((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
		CComVariant v(1);
		CComVariant v2(2);
		CComVariant v3(_T(""));
		//下面的Replace 对应的替换的范围是哪里.
		// 1 代表一个 2 代表整个文档
		//myFind.Execute(Text,vFalse,vFalse,vFalse,vFalse,vFalse,vTrue,&v,vFalse,re,&v2,vOpt,vOpt,vOpt,vOpt);
		myFind.Execute(Text, vFalse, vFalse, vFalse, vFalse, vFalse,
			vTrue, &v, vFalse, &re, &v2, vOpt, vOpt, vOpt, vOpt);
	}
}
void CmyWord::GetWordText(CString &Text){
	//CComVariant vOpt;
	COleVariant vOpt((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	doc = app.GetActiveDocument();//获得当前激活文档 就是当前正在编辑文档
	if (!doc.m_lpDispatch){
		AfxMessageBox(_T("获取激活文档对象失败"));
		return;
	}
	else{
		range = doc.Range(vOpt, vOpt);
		Text = range.GetText();
		AfxMessageBox(Text);
	}
}
//打印代码我直接Cppy 别人的 因为我没有打印机所以不好做测试
//这里只是为了方便大家
void CmyWord::PrintWord(){
	doc = app.GetActiveDocument();
	if (!doc.m_lpDispatch){
		AfxMessageBox(_T("获取激活文档对象失败"));
		return;
	}
	else{
		COleVariant covTrue((short)TRUE),
			covFalse((short)FALSE),
			covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
		doc.PrintOut(covFalse,              // Background.
			covOptional,           // Append.
			covOptional,           // Range.
			covOptional,           // OutputFileName.
			covOptional,           // From.
			covOptional,           // To.
			covOptional,           // Item.
			COleVariant((long)1),  // Copies.
			covOptional,           // Pages.
			covOptional,           // PageType.
			covOptional,           // PrintToFile.
			covOptional,           // Collate.
			covOptional,           // ActivePrinterMacGX.
			covOptional,           // ManualDuplexPrint.
			covOptional,           // PrintZoomColumn  New with Word 2002
			covOptional,           // PrintZoomRow          ditto
			covOptional,           // PrintZoomPaperWidth   ditto
			covOptional);          // PrintZoomPaperHeight  ditto*/
	}
}
void CmyWord::AppClose(){
	COleVariant vOpt((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	if (!app.m_lpDispatch){
		AfxMessageBox(_T("获取Word 对象失败,关闭操作失败"));
		return;
	}
	else{
		app.Quit(vOpt, vOpt, vOpt);
		//这里释放资源好像不是很好，所以我就在析构函数去处理了。
	}
}