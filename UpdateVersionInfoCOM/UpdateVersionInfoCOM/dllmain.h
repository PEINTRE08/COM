// dllmain.h: 模块类的声明。

class CUpdateVersionInfoCOMModule : public ATL::CAtlDllModuleT< CUpdateVersionInfoCOMModule >
{
public :
	DECLARE_LIBID(LIBID_UpdateVersionInfoCOMLib)
	DECLARE_REGISTRY_APPID_RESOURCEID(IDR_UPDATEVERSIONINFOCOM, "{545442FC-463D-4320-A259-226FE5DD06F4}")
};

extern class CUpdateVersionInfoCOMModule _AtlModule;
