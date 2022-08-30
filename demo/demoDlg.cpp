
// demoDlg.cpp: 实现文件
//

#include "pch.h"
#include "framework.h"
#include "demo.h"
#include "demoDlg.h"
#include "afxdialogex.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// 用于应用程序“关于”菜单项的 CAboutDlg 对话框

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// 对话框数据
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_ABOUTBOX };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

// 实现
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialogEx(IDD_ABOUTBOX)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()


// CdemoDlg 对话框


CdemoDlg::CdemoDlg(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_DEMO_DIALOG, pParent),
	m_isImport(FALSE),
	m_number(0)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CdemoDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_LIST2, m_list_show);
}

BEGIN_MESSAGE_MAP(CdemoDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_BUTTON1, &CdemoDlg::OnBnClickedButton1)
	ON_BN_CLICKED(IDOK, &CdemoDlg::OnBnClickedOk)
	ON_BN_CLICKED(IDC_BUTTON2, &CdemoDlg::OnBnClickedButton2)
	ON_BN_CLICKED(IDC_BUTTON3, &CdemoDlg::OnBnClickedButton3)
	ON_BN_CLICKED(IDC_BUTTON4, &CdemoDlg::OnBnClickedButton4)
END_MESSAGE_MAP()


// CdemoDlg 消息处理程序

BOOL CdemoDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// 将“关于...”菜单项添加到系统菜单中。

	// IDM_ABOUTBOX 必须在系统命令范围内。
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != nullptr)
	{
		BOOL bNameValid;
		CString strAboutMenu;
		bNameValid = strAboutMenu.LoadString(IDS_ABOUTBOX);
		ASSERT(bNameValid);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// 设置此对话框的图标。  当应用程序主窗口不是对话框时，框架将自动
	//  执行此操作
	SetIcon(m_hIcon, TRUE);			// 设置大图标
	SetIcon(m_hIcon, FALSE);		// 设置小图标

	// TODO: 在此添加额外的初始化代码

	return TRUE;  // 除非将焦点设置到控件，否则返回 TRUE
}

void CdemoDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialogEx::OnSysCommand(nID, lParam);
	}
}

// 如果向对话框添加最小化按钮，则需要下面的代码
//  来绘制该图标。  对于使用文档/视图模型的 MFC 应用程序，
//  这将由框架自动完成。

void CdemoDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // 用于绘制的设备上下文

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// 使图标在工作区矩形中居中
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// 绘制图标
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

//当用户拖动最小化窗口时系统调用此函数取得光标
//显示。
HCURSOR CdemoDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}



void CdemoDlg::OnBnClickedButton1()
{
	if (m_student.size() == 0) {
		m_list_show.InsertString(m_number++, "内存中无学生信息,请先读取学生信息");
		return;
	}
	char* buffer = "../testData";
	CmyWord ctrlWord;
	if (ctrlWord.Create()) {
		CString name;
		name = "student.docx";
		CString pt = "\\" + name;
		CString docxAP = buffer + pt;
		ctrlWord.Open(docxAP, FALSE, FALSE);
		m_list_show.InsertString(m_number++, "在word中新建表格");
		ctrlWord.CreateTable(int (m_student.size() + 1), 4);
		ctrlWord.WriteCellText(1, 1, "学号");
		ctrlWord.WriteCellText(1, 2, "姓名");
		ctrlWord.WriteCellText(1, 3, "电话号码");
		ctrlWord.WriteCellText(1, 4, "籍贯");
		m_list_show.InsertString(m_number++, "正在向表格中输入数据");
		for (int i = 0; i < m_student.size(); i++) {
			ctrlWord.WriteCellText(i + 2, 1, m_student[i].Code);
			ctrlWord.WriteCellText(i + 2, 2, m_student[i].Name);
			ctrlWord.WriteCellText(i + 2, 3, m_student[i].PhoneCode);
			ctrlWord.WriteCellText(i + 2, 4, m_student[i].Native);
		}
		m_list_show.InsertString(m_number++, "数据输入完成");
		ctrlWord.Save();
		m_list_show.InsertString(m_number++, "文件已保存");
	}
	// TODO: 在此添加控件通知处理程序代码
}


void CdemoDlg::OnBnClickedOk()
{
	// TODO: 在此添加控件通知处理程序代码
	string databaseName;
	databaseName = "F:\\officeApiWrapper_vs2010\\testdata\student.accdb";
	m_connect.setbaseName(databaseName);
	m_list_show.InsertString(m_number++, "正在连接数据库");
	if (m_connect.InitADOaccess()) {
		m_list_show.InsertString(m_number++, "成功连接数据库");
		_bstr_t sql = "create table student(code NTEXT, name NTEXT, phoneNumber NTEXT, native NTEXT)";
		if (m_connect.ExecuteSQL(sql))
			m_list_show.InsertString(m_number++, "成功新建student表");
		else
			m_list_show.InsertString(m_number++, "建表失败,该表可能已存在");
	}
	else
		m_list_show.InsertString(m_number++, "连接数据库失败,请重新连接");
	//	CDialogEx::OnOK();
}


void CdemoDlg::OnBnClickedButton2()
{
	// TODO: 在此添加控件通知处理程序代码
	if (m_isImport)return;
	string name;
	name = "../testData/student.txt";
	ifstream inFile;
	inFile.open(name);
	m_list_show.InsertString(m_number++, "正在从student.txt中读取数据导入至数据库");
	while (inFile.good()) {
		string code;
		string name;
		string phone;
		string native;
		StuInf stu;
		inFile >> code >> name >> phone >> native;
		stu.Code = code.c_str();
		stu.Name = name.c_str();
		stu.PhoneCode = phone.c_str();
		stu.Native = native.c_str();
		if (!m_connect.addMessage(stu)) {
			CString pt;
			pt = "该条数据导入失败: " + stu.Code + " , " + stu.Name + " , " + stu.PhoneCode + " ," + stu.Native + " . ";
			m_list_show.InsertString(m_number++, pt);
			break;
		}
	}
	m_list_show.InsertString(m_number++, "数据导入完成");
}


void CdemoDlg::OnBnClickedButton3()
{
	// TODO: 在此添加控件通知处理程序代码
	m_student = m_connect.getAlldata();
	if (m_student.size() != 0)
		m_list_show.InsertString(m_number++, "数据读取成功");
	else
		m_list_show.InsertString(m_number++, "数据读取失败");
}


void CdemoDlg::OnBnClickedButton4()
{
	// TODO: 在此添加控件通知处理程序代码
	_bstr_t sql = "drop table student";
	if (m_connect.ExecuteSQL(sql)) {
		m_list_show.InsertString(m_number++, "student表删除成功");
	}
	else {
		m_list_show.InsertString(m_number++, "student表删除失败,student表可能不存在");
	}
}
