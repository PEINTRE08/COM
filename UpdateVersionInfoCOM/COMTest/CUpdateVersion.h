// �����Ϳ������á�����ࡱ�����ļ�������ɵ� IDispatch ��װ����

#import "D:\\VS2015\\UpdateVersionInfoCOM\\Debug\\UpdateVersionInfoCOM.dll" no_namespace
// CUpdateVersion ��װ����

class CUpdateVersion : public COleDispatchDriver
{
public:
	CUpdateVersion() {} // ���� COleDispatchDriver Ĭ�Ϲ��캯��
	CUpdateVersion(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CUpdateVersion(const CUpdateVersion& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// ����
public:

	// ����
public:


	// IUpdateVersionInfo ����
public:
	long UpdatePEInfo(LPCTSTR lpFile, LPCTSTR lpKey, BSTR * lppValue)
	{
		long result;
		static BYTE parms[] = VTS_BSTR VTS_BSTR VTS_PBSTR;
		InvokeHelper(0x1, DISPATCH_METHOD, VT_I4, (void*)&result, parms, lpFile, lpKey, lppValue);
		return result;
	}
	long UpdatePEVersion(LPCTSTR lpFile, LPCTSTR lpKey, LPCTSTR lpValue)
	{
		long result;
		static BYTE parms[] = VTS_BSTR VTS_BSTR VTS_BSTR;
		InvokeHelper(0x2, DISPATCH_METHOD, VT_I4, (void*)&result, parms, lpFile, lpKey, lpValue);
		return result;
	}

	// IUpdateVersionInfo ����
public:

};
