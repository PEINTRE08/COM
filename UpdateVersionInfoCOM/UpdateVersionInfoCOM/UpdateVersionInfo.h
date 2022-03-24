// UpdateVersionInfo.h : CUpdateVersionInfo ������

#pragma once
#include "resource.h"       // ������
#include <map>
#include <vector>


#include "UpdateVersionInfoCOM_i.h"

#pragma comment(lib, "version.lib")

#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Windows CE ƽ̨(�粻�ṩ��ȫ DCOM ֧�ֵ� Windows Mobile ƽ̨)���޷���ȷ֧�ֵ��߳� COM ���󡣶��� _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA ��ǿ�� ATL ֧�ִ������߳� COM ����ʵ�ֲ�����ʹ���䵥�߳� COM ����ʵ�֡�rgs �ļ��е��߳�ģ���ѱ�����Ϊ��Free����ԭ���Ǹ�ģ���Ƿ� DCOM Windows CE ƽ̨֧�ֵ�Ψһ�߳�ģ�͡�"
#endif

using namespace ATL;

#define roundoffs(a,b,r) (((BYTE *) (b) - (BYTE *) (a) + ((r) - 1)) & ~((r) - 1))
#define roundpos(a,b,r) (((BYTE *) (a)) + roundoffs(a,b,r))
#define SIGNATURE 0xFEEF04BD
template<typename T>
inline T Round(T value, int modula = 4) {
	return value + ((value % modula > 0) ? (modula - value % modula) : 0);
}

typedef std::pair<std::wstring, std::wstring> VersionString;
typedef std::pair<const BYTE* const, const size_t> OffsetLengthPair;

#pragma pack(push,1)
typedef struct _VS_VERSION_HEADER {
	WORD wLength;
	WORD wValueLength;
	WORD wType;
} VS_VERSION_HEADER;
#pragma pack(pop)

#pragma pack(push,1)
typedef struct {
	VS_VERSION_HEADER Header;
	WCHAR            szKey[16];
	WORD             Padding1[1];
	VS_FIXEDFILEINFO Value;
	WORD             Padding2[1];
	WORD             Children[1];
} VS_VERSIONINFO;
#pragma pack(pop)

#pragma pack(push,1)
typedef struct {
	VS_VERSION_HEADER Header;
	WCHAR            szKey[1];
}VS_VERSION_STRING;
#pragma pack(pop)

//#pragma pack(push,1)
//typedef struct {
//	VS_VERSION_HEADER Header;
//	WCHAR       szKey[1];
//	WORD        Padding[1];
//	StringTable Children[1];
//} StringFileInfo;
//#pragma pack(pop)
//
//#pragma pack(push,1)
//typedef struct {
//	VS_VERSION_HEADER Header;
//	WCHAR  szKey[1];
//	WORD   Padding[1];
//	String Children[1];
//} StringTable;
//#pragma pack(pop)
//
//#pragma pack(push,1)
//typedef struct {
//	VS_VERSION_HEADER Header;
//	WCHAR szKey[1];
//	WORD  Padding[1];
//	WORD  Value[1];
//} String;
//#pragma pack(pop)

typedef struct LANGANDCODEPAGE {
	WORD wLanguage;
	WORD wCodePage;
} TRANSLATE;

struct VersionStringTable {
	TRANSLATE encoding;
	std::vector<VersionString> strings;
};

struct VersionStampValue {
	WORD valueLength = 0; // stringfileinfo, stringtable: 0; string: Value size in WORD; var: Value size in bytes
	WORD type = 0; // 0: binary data; 1: text data
	std::wstring key; // stringtable: 8-digit hex stored as UTF-16 (hiword: hi6: sublang, lo10: majorlang; loword: code page); must include zero words to align next member on 32-bit boundary
	std::vector<BYTE> value; // string: zero-terminated string; var: array of language & code page ID pairs
	std::vector<VersionStampValue> children;

	size_t GetLength() const;
	std::vector<BYTE> Serialize() const;
};

// CUpdateVersionInfo

class ATL_NO_VTABLE CUpdateVersionInfo :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CUpdateVersionInfo, &CLSID_UpdateVersionInfo>,
	public IDispatchImpl<IUpdateVersionInfo, &IID_IUpdateVersionInfo, &LIBID_UpdateVersionInfoCOMLib, /*wMajor =*/ 1, /*wMinor =*/ 0>
{
public:
	CUpdateVersionInfo()
	{
	}

DECLARE_REGISTRY_RESOURCEID(IDR_UPDATEVERSIONINFO)


BEGIN_COM_MAP(CUpdateVersionInfo)
	COM_INTERFACE_ENTRY(IUpdateVersionInfo)
	COM_INTERFACE_ENTRY(IDispatch)
END_COM_MAP()



	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}

	void FinalRelease()
	{
	}

public:



	STDMETHOD(UpdatePEInfo)(BSTR lpFile, BSTR lpKey, BSTR* lppValue, LONG* lppResult);
	BOOL UpdatePEVersion(CString strFile, CString strCode, CString strValue);
	void DeserializeVarFileInfo(const unsigned char * offset, std::vector<TRANSLATE>& translations);
	void DeserializeVersionStringFileInfo(const BYTE * offset, size_t length, std::vector<VersionStringTable>& stringTables);
	VersionStringTable DeserializeVersionStringTable(const BYTE * tableData);
	OffsetLengthPair GetChildrenData(const BYTE * entryData);
	BOOL ApiUpdatePEVersion(CString strFile, CString strCode, CString strValue);
	CStringA CStringTochar(CString str1);
	BOOL SplitString(CStringArray * pArray, CString strMultiValue, CString strSymbol);
	BOOL GetApiPEInfo(CString strFile, CString strCode, CString & strValue);
	STDMETHOD(UpdatePEVersion)(BSTR lpFile, BSTR lpKey, BSTR lpValue, LONG* lppResult);
};

OBJECT_ENTRY_AUTO(__uuidof(UpdateVersionInfo), CUpdateVersionInfo)
