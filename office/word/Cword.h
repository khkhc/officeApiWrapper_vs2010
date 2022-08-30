#pragma once

#ifndef CMYWORD_H

#define CMYWORD_H

#include "msword.h"
#include <ATLBASE.H>
//段落对齐的属性
enum Alignment { wdAlignParagraphCenter = 1, wdAlignParagraphRight, wdAlignParagraphJustify };
enum SaveType {
	wdFormatDocument = 0,
	wdFormatWebArchive = 9,
	wdFormatHTML = 8,
	wdFormatFilteredHTML = 10,
	wdFormatTemplate = 1

};
class CmyWord{
	//一些对象申明
public:
	_Application app;//创建word
	Documents docs;//word文档集合
	_Document doc;//一个word文件
	_Font font;//字体对象
	Selection sel;//选择编辑对象 没有对象的时候就是插入点
	Table tab;//表格对象
	Range range;

public:
	CmyWord();//构造函数
	virtual ~CmyWord();//析构函数
	void ShowApp(BOOL flag);
	void AppClose();
	BOOL InitCOM();//对COM进行初始化
	BOOL CreateAPP();//创建一个word程序
	BOOL CreateDocument();//创建word文档
	BOOL Create();//创建一个word程序和Word文档


	BOOL Open(CString FileName, BOOL ReadOnly = FALSE, BOOL  AddToRecentFiles = FALSE);//打开一个word文档;
	BOOL Close(BOOL SaveChange = FALSE);//关闭一个word文档
	BOOL Save();//保存文档
	BOOL SaveAs(CString FileName, int SaveType = 0);//保存类型

	//////////////////////////文件写操作操作/////////////////////////////////////////////

	void WriteText(CString Text);//写入文本
	void NewLine(int nCount = 1);//回车换N行
	void WriteTextNewLineText(CString Text, int nCount = 1);//回测换N行写入文字

	//////////////////////////////////////////////////////////////////////////
	   
	//////////////////////////字体设置////////////////////////////////////////
	void SetFont(CString FontName, int FontSize = 9, long FontColor = 0, long FontBackColor = 0);
	void SetFont(BOOL Blod, BOOL Italic = FALSE, BOOL UnderLine = FALSE);
	void SetTableFont(int Row, int Column, CString FontName, int FontSize = 9, long FontColor = 0, long FontBackColor = 0);
	//void SetTableFont();//统一对表格的文字做出处理.

		/////////////////////////表格操作/////////////////////////////////////
	void CreateTable(int Row, int Column);
	void WriteCellText(int Row, int Column, CString Text);
	/////////////////////////////设置对齐属性///////////////////////////////////////
	void SetParaphformat(int Alignment);

	/////////////////////////////一些常用操作///////////////////////////////////////
	//查找字符串 然后全部替换
	void FindWord(CString FindW, CString RelWord);
	//获取Word 纯文本内容
	void GetWordText(CString &Text);
	//Word 打印
	void PrintWord();
};
#endif


