
// demoDlg.h: 头文件
//

#pragma once


// CdemoDlg 对话框
class CdemoDlg : public CDialogEx
{
	// 构造
public:
	CdemoDlg(CWnd* pParent = nullptr);	// 标准构造函数

	// 对话框数据
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_DEMO_DIALOG	};
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 支持


	// 实现
protected:
	HICON m_hIcon;

	// 生成的消息映射函数
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedButton1();
	// 用于显示执行命令中的过程
	CListBox m_list_show;
	//用于数据库的连接
	ADOaccess m_connect;
	//储存从数据库中读取的数据
	vector<StuInf> m_student;
	//判断是否已经导入数据
	BOOL m_isImport;
	//记录已执行命令数
	int m_number;
	afx_msg void OnBnClickedOk();
	afx_msg void OnBnClickedButton2();
	afx_msg void OnBnClickedButton3();
	afx_msg void OnBnClickedButton4();
};
