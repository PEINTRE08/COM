// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装器类

#import "D:\\VS2015\\UpdateVersionInfoCOM\\Debug\\UpdateVersionInfoCOM.dll" no_namespace
// CUpdateVersion 包装器类

class CUpdateVersion : public COleDispatchDriver
{
public:
	CUpdateVersion() {} // 调用 COleDispatchDriver 默认构造函数
	CUpdateVersion(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CUpdateVersion(const CUpdateVersion& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 特性
public:

	// 操作
public:


	// IUpdateVersionInfo 方法
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

	// IUpdateVersionInfo 属性
public:

};
