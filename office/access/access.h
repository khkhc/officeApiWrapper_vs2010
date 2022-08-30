#pragma once
#import ".\msado15.dll" \
no_namespace \
rename("EOF", "adoEOF") rename("DataTypeEnum", "adoDataTypeEnum") \
rename("FieldAttributeEnum", "adoFielAttributeEnum") rename("EditModeEnum", "adoEditModeEnum") \
rename("LockTypeEnum", "adoLockTypeEnum") rename("RecordStatusEnum", "adoRecordStatusEnum") \
rename("ParameterDirectionEnum", "adoParameterDirectionEnum") \
rename("Field", "adoField") rename("Fields", "adoFields")

#include <string>
#include <vector>
using namespace std;
CString VariantToString(VARIANT var);

struct StuInf {
	CString Code;
	CString Name;
	CString PhoneCode;
	CString Native;
};
class ADOaccess {
public:
	ADOaccess() {}
	ADOaccess(string dataname):DataBaseName(dataname){}

	BOOL InitADOaccess(string DatabaseName);		//连接access数据库,初始化m_pConnection

	BOOL InitADOaccess() {
		return InitADOaccess(DataBaseName);
	}

	void ExitAccess();								//关闭数据库连接

	_RecordsetPtr& GetRecordSet(_bstr_t bstrSQL);	//获取记录集

	BOOL ExecuteSQL(_bstr_t bstrSQL);				//执行数据库语句

	bool addMessage(StuInf pt);                     //添加信息

	void setbaseName(string name) {					//设置所要连接的数据库名称	
		DataBaseName = name;
	}					

	vector<StuInf> getAlldata();

private:
	_ConnectionPtr m_pConnection;                   //连接access数据库的链接对象
	_RecordsetPtr m_pRecordset;                     //结果集对象
	string DataBaseName;							//数据库名称
};
