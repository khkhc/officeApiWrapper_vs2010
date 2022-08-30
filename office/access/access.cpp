#include "afxole.h"
#include "atlstr.h"
#include "atltime.h"
#include "access.h"

BOOL ADOaccess::InitADOaccess(string DatabaseName)
{
	CoInitialize(NULL);         //初始化OLE/COM库环境  
	try {
		m_pConnection.CreateInstance("ADODB.Connection");  //创建连接对象实例
		CString connect1 = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=";
		CString Connect = connect1 + DatabaseName.c_str();
		_bstr_t strConnect = Connect;
		m_pConnection->Open(strConnect, "", "", adModeUnknown); //打开数据库
	}
	catch (_com_error e) {
		return FALSE;
	}
	return TRUE;
}

void ADOaccess::ExitAccess()
{
	if (m_pRecordset != NULL){
		m_pRecordset->Close();
	}
	m_pConnection->Close();
	//释放环境
	CoUninitialize();
}

_RecordsetPtr & ADOaccess::GetRecordSet(_bstr_t bstrSQL)
{
	try {
		if (m_pConnection == NULL) //判断Connection对象是否为空	
			InitADOaccess(DataBaseName); //如果为空则重新连接数据库		
		m_pRecordset.CreateInstance("ADODB.Recordset"); //创建记录集对象		
		//获取数据表中的数据		
		m_pRecordset->Open(bstrSQL, m_pConnection.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);
	}	
	catch (_com_error e) {	//捕获异常
		_RecordsetPtr pt;
		return pt;
	}	
	return m_pRecordset;
	// TODO: 在此处插入 return 语句
}

BOOL ADOaccess::ExecuteSQL(_bstr_t bstrSQL)
{
	try {
		if(!m_pConnection)
			InitADOaccess(DataBaseName); //如果为空则重新连接数据库		
		m_pConnection->Execute(bstrSQL, NULL , adCmdText);//执行sql命令
	}
	catch (_com_error e) {
		return FALSE;
	}
	return TRUE;
}

bool ADOaccess::addMessage(StuInf pt)
{
	if (!m_pConnection)
		InitADOaccess(DataBaseName); //如果为空则重新连接数据库	
	_variant_t RecordsAffected;                        //定义插入对象
	CString AddSql;
	CString median = "INSERT INTO student(code,name,phoneNumber,native) VALUES('";
	AddSql = median + pt.Code + "','" + pt.Name + "','" + pt.PhoneCode + "','" + pt.Native + "')";
	return ExecuteSQL(_bstr_t(AddSql));
}

vector<StuInf> ADOaccess::getAlldata()
{
	vector<StuInf> alldata;
	StuInf data;
	if (!m_pConnection)
		InitADOaccess(DataBaseName); //如果为空则重新连接数据库	
	_bstr_t bstrSQL = "select * from student";
	m_pRecordset = GetRecordSet(bstrSQL);
	if (m_pRecordset) {
		while (!m_pRecordset->adoEOF) {			
			data.Code = VariantToString(m_pRecordset->GetadoFields()->GetItem(long(0))->Value);
			data.Name = VariantToString(m_pRecordset->GetadoFields()->GetItem(long(1))->Value);
			data.PhoneCode = VariantToString(m_pRecordset->GetadoFields()->GetItem(long(2))->Value);
			data.Native = VariantToString(m_pRecordset->GetadoFields()->GetItem(long(3))->Value);
			
			alldata.push_back(data);
			m_pRecordset->MoveNext();
		}
	}
	m_pRecordset = _RecordsetPtr();
	return alldata;
}

CString VariantToString(VARIANT var) {
	CString strValue;
	_variant_t var_t;
	_bstr_t bstr_t1;
	time_t cur_time;
	CTime time_value;
	COleCurrency var_currency;
	switch (var.vt)
	{
	case VT_EMPTY:
	case VT_NULL:strValue = _T(""); break;
	case VT_UI1:strValue.Format(_T("%d"), var.bVal); break;
	case VT_I2:strValue.Format(_T("%d"), var.iVal); break;
	case VT_I4:strValue.Format(_T("%d"), var.lVal); break;
	case VT_R4:strValue.Format(_T("%f"), var.fltVal); break;
	case VT_R8:strValue.Format(_T("%f"), var.dblVal); break;
	case VT_CY:
		var_currency = var;
		strValue = var_currency.Format(0); break;
	case VT_BSTR:
		var_t = var;
		bstr_t1 = var_t;
		strValue.Format(_T("%s"), (const char *)bstr_t1); break;
	case VT_DATE:
		cur_time = var.date;
		time_value = cur_time;
		strValue.Format(_T("%A,%B,%d,%Y")); break;
	case VT_BOOL:strValue.Format(_T("%d"), var.boolVal); break;
	default:strValue = _T(""); break;
	}
	return strValue;
}
