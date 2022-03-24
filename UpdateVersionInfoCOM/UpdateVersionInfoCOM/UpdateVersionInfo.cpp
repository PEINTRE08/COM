// UpdateVersionInfo.cpp : CUpdateVersionInfo 的实现

#include "stdafx.h"
#include "UpdateVersionInfo.h"
#include <sstream> // wstringstream
#include <iomanip> // setw, setfill
// CUpdateVersionInfo
LANGID klangCN = 2052;
LANGID kCodePageCN = 1200;


STDMETHODIMP CUpdateVersionInfo::UpdatePEInfo(BSTR lpFile, BSTR lpKey, BSTR* lppValue, LONG* lppResult)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// TODO: 在此添加实现代码
	CString strFile, strCode, strValue;

	strFile = lpFile;
	strCode = lpKey;

	BOOL rc = this->GetApiPEInfo(strFile, strCode, strValue);
	if (rc == FALSE)
	{
		*lppResult = FALSE;
		return S_FALSE;
	}
	else
	{
		*lppResult = TRUE;
		*lppValue = strValue.AllocSysString();
	}
	return S_OK;
}

BOOL CUpdateVersionInfo::UpdatePEVersion(CString strFile, CString strCode, CString strValue)
{
	HMODULE hModule = LoadLibraryExW(strFile, NULL, DONT_RESOLVE_DLL_REFERENCES | LOAD_LIBRARY_AS_DATAFILE);
	if (hModule == NULL)
		return FALSE;

	HRSRC hRsrc = FindResourceExW(hModule, RT_VERSION, MAKEINTRESOURCEW(1), klangCN);
	if (hRsrc == NULL) 
		return FALSE;
	
	HGLOBAL hGlobal = LoadResource(hModule, hRsrc);
	if (hGlobal == NULL) 
		return FALSE;
	
	void* pData = LockResource(hGlobal);
	if (pData == NULL)
		return FALSE;
	
	DWORD size = SizeofResource(hModule, hRsrc);
	if (size == 0) 
		return FALSE;

	// 获取 VS_FIXEDFILEINFO
	auto pVersionInfo = reinterpret_cast<const VS_VERSIONINFO*>(pData);
	VS_FIXEDFILEINFO pFixedFileInfo;
	if (pVersionInfo->Header.wValueLength > 0)
		pFixedFileInfo = pVersionInfo->Value;

	if (pFixedFileInfo.dwSignature != SIGNATURE)
	{
		pFixedFileInfo = { 0 };
		pFixedFileInfo.dwSignature = SIGNATURE;
		pFixedFileInfo.dwFileType = VFT_APP;
	}
	/////////////////////////

	// 确定 Value or Children 的位置
	const BYTE* fixedFileInfoEndOffset = reinterpret_cast<const BYTE*>(&pVersionInfo->szKey) + (wcslen(pVersionInfo->szKey) + 1) * sizeof(WCHAR) + pVersionInfo->Header.wValueLength;
	const BYTE* pVersionInfoChildren = reinterpret_cast<const BYTE*>(Round(reinterpret_cast<ptrdiff_t>(fixedFileInfoEndOffset)));
	size_t versionInfoChildrenOffset = pVersionInfoChildren - pData;
	size_t versionInfoChildrenSize = pVersionInfo->Header.wLength - versionInfoChildrenOffset;

	const auto childrenEndOffset = pVersionInfoChildren + versionInfoChildrenSize;
	const auto resourceEndOffset = static_cast<BYTE*>(pData) + size;
	/////////////////////////////////

	// 获取 StringFileInfo and VarFileInfo 
	std::vector<VersionStringTable> stringTables;
	std::vector<TRANSLATE> supportedTranslations;

	for (auto p = pVersionInfoChildren; p < childrenEndOffset && p < resourceEndOffset;) {
		auto pKey = reinterpret_cast<const VS_VERSION_STRING*>(p)->szKey;
		auto versionInfoChildData = GetChildrenData(p);
		if (wcscmp(pKey, L"StringFileInfo") == 0) {
			DeserializeVersionStringFileInfo(versionInfoChildData.first, versionInfoChildData.second, stringTables);
		}
		else if (wcscmp(pKey, L"VarFileInfo") == 0) {
			DeserializeVarFileInfo(versionInfoChildData.first, supportedTranslations);
		}
		p += Round(reinterpret_cast<const VS_VERSION_STRING*>(p)->Header.wLength);
	}

	if (stringTables.empty())
	{
		TRANSLATE translate = { klangCN, kCodePageCN };
		stringTables.push_back({ translate });
		supportedTranslations.push_back(translate);
	}
	////////////////////////////////////////////

	if (strCode == "ProductVersion" || strCode == "FileVersion")
	{
		unsigned short ver[4] = { 0 };
		CStringArray sArray;
		this->SplitString(&sArray, strValue, _T("."));
		for (size_t i = 0; i < sArray.GetSize(); i++)
		{
			CStringA s = CStringTochar(sArray[i]);
			ver[i] = atoi(s);
		}
		if (strCode == "ProductVersion")
		{
			pFixedFileInfo.dwProductVersionMS = ver[0] << 16 | ver[1];
			pFixedFileInfo.dwProductVersionLS = ver[2] << 16 | ver[3];
		}
		if (strCode == "FileVersion")
		{
			pFixedFileInfo.dwFileVersionMS = ver[0] << 16 | ver[1];
			pFixedFileInfo.dwFileVersionLS = ver[2] << 16 | ver[3];
		}
	}

	for (auto j = stringTables.begin(); j != stringTables.end(); ++j) 
	{
		auto& stringPairs = j->strings;
		for (auto k = stringPairs.begin(); k != stringPairs.end(); ++k) 
		{
			if (k->first == strCode.AllocSysString())
			{
				k->second = strValue;
				goto stop;
			}
		}
		// Not found, append one for all tables.
		stringPairs.push_back(VersionString(strCode, strValue));
	}

stop:

	FreeLibrary(hModule);
	hModule = NULL;

	// 重组数据
	HANDLE hResource = BeginUpdateResource(strFile.AllocSysString(), FALSE);
	if (NULL == hResource)
		return FALSE;

	VersionStampValue versionInfo;
	versionInfo.key = L"VS_VERSION_INFO";
	versionInfo.type = 0;
	auto fixedsize = sizeof(VS_FIXEDFILEINFO);
	versionInfo.valueLength = fixedsize;

	auto& dst = versionInfo.value;
	dst.resize(fixedsize);

	memcpy(&dst[0], &pFixedFileInfo, fixedsize);

	{
		VersionStampValue stringFileInfo;
		stringFileInfo.key = L"StringFileInfo";
		stringFileInfo.type = 1;
		stringFileInfo.valueLength = 0;

		for (const auto& iTable : stringTables) {
			VersionStampValue stringTableRaw;
			stringTableRaw.type = 1;
			stringTableRaw.valueLength = 0;

			{
				auto& translate = iTable.encoding;
				std::wstringstream ss;
				ss << std::hex << std::setw(8) << std::setfill(L'0') << (translate.wLanguage << 16 | translate.wCodePage);
				stringTableRaw.key = ss.str();
			}

			for (const auto& iString : iTable.strings) {
				const auto& stringValue = iString.second;
				auto strLenNullTerminated = stringValue.length() + 1;

				VersionStampValue stringRaw;
				stringRaw.type = 1;
				stringRaw.key = iString.first;
				stringRaw.valueLength = strLenNullTerminated;

				auto size = strLenNullTerminated * sizeof(WCHAR);
				auto& dst = stringRaw.value;
				dst.resize(size);

				auto src = stringValue.c_str();

				memcpy(&dst[0], src, size);

				stringTableRaw.children.push_back(std::move(stringRaw));
			}

			stringFileInfo.children.push_back(std::move(stringTableRaw));
		}

		versionInfo.children.push_back(std::move(stringFileInfo));
	}

	{
		VersionStampValue varFileInfo;
		varFileInfo.key = L"VarFileInfo";
		varFileInfo.type = 1;
		varFileInfo.valueLength = 0;

		{
			VersionStampValue varRaw;
			varRaw.key = L"Translation";
			varRaw.type = 0;

			{
				auto newValueSize = sizeof(DWORD);
				auto& dst = varRaw.value;
				dst.resize(supportedTranslations.size() * newValueSize);

				for (auto iVar = 0; iVar < supportedTranslations.size(); ++iVar) {
					auto& translate = supportedTranslations[iVar];
					auto var = DWORD(translate.wCodePage) << 16 | translate.wLanguage;
					memcpy(&dst[iVar * newValueSize], &var, newValueSize);
				}

				varRaw.valueLength = varRaw.value.size();
			}

			varFileInfo.children.push_back(std::move(varRaw));
		}

		versionInfo.children.push_back(std::move(varFileInfo));
	}

	std::vector<BYTE> out = std::move(versionInfo.Serialize());

	// 更新数据
	if (!UpdateResourceW(hResource, RT_VERSION, MAKEINTRESOURCEW(1), klangCN,
		&out[0], static_cast<DWORD>(out.size()))) {
		return FALSE;
	}

	BOOL bResult = EndUpdateResourceW(hResource, FALSE);
	return bResult ? TRUE : FALSE;
}

size_t VersionStampValue::GetLength() const {
	size_t bytes = sizeof(VS_VERSION_HEADER);
	bytes += static_cast<size_t>(key.length() + 1) * sizeof(WCHAR);
	if (!value.empty())
		bytes = Round(bytes) + value.size();
	for (const auto& child : children)
		bytes = Round(bytes) + static_cast<size_t>(child.GetLength());
	return bytes;
}

std::vector<BYTE> VersionStampValue::Serialize() const {
	std::vector<BYTE> data = std::vector<BYTE>(GetLength());

	size_t offset = 0;

	VS_VERSION_HEADER header = { static_cast<WORD>(data.size()), valueLength, type };
	memcpy(&data[offset], &header, sizeof(header));
	offset += sizeof(header);

	auto keySize = static_cast<size_t>(key.length() + 1) * sizeof(WCHAR);
	memcpy(&data[offset], key.c_str(), keySize);
	offset += keySize;

	if (!value.empty()) {
		offset = Round(offset);
		memcpy(&data[offset], &value[0], value.size());
		offset += value.size();
	}

	for (const auto& child : children) {
		offset = Round(offset);
		size_t childLength = child.GetLength();
		std::vector<BYTE> src = child.Serialize();
		memcpy(&data[offset], &src[0], childLength);
		offset += childLength;
	}

	return std::move(data);
}

void CUpdateVersionInfo::DeserializeVarFileInfo(const unsigned char* offset, std::vector<TRANSLATE>& translations) {
	const auto translatePairs = GetChildrenData(offset);

	const auto top = reinterpret_cast<const DWORD* const>(translatePairs.first);
	for (auto pTranslatePair = top; pTranslatePair < top + translatePairs.second; pTranslatePair += sizeof(DWORD)) {
		auto codePageLangIdPair = *pTranslatePair;
		TRANSLATE translate;
		translate.wLanguage = codePageLangIdPair;
		translate.wCodePage = codePageLangIdPair >> 16;
		translations.push_back(translate);
	}
}

void CUpdateVersionInfo::DeserializeVersionStringFileInfo(const BYTE* offset, size_t length, std::vector<VersionStringTable>& stringTables) {
	for (auto posStringTables = 0U; posStringTables < length;) {
		auto stringTableEntry = DeserializeVersionStringTable(offset + posStringTables);
		stringTables.push_back(stringTableEntry);
		posStringTables += Round(reinterpret_cast<const VS_VERSION_STRING*>(offset + posStringTables)->Header.wLength);
	}
}

VersionStringTable CUpdateVersionInfo::DeserializeVersionStringTable(const BYTE* tableData) {
	auto strings = GetChildrenData(tableData);
	auto stringTable = reinterpret_cast<const VS_VERSION_STRING*>(tableData);
	auto end_ptr = const_cast<WCHAR*>(stringTable->szKey + (8 * sizeof(WCHAR)));
	auto langIdCodePagePair = static_cast<DWORD>(wcstol(stringTable->szKey, &end_ptr, 16));

	VersionStringTable tableEntry;

	// unicode string of 8 hex digits
	tableEntry.encoding.wLanguage = langIdCodePagePair >> 16;
	tableEntry.encoding.wCodePage = langIdCodePagePair;

	for (auto posStrings = 0U; posStrings < strings.second;) {
		const auto stringEntry = reinterpret_cast<const VS_VERSION_STRING* const>(strings.first + posStrings);
		const auto stringData = GetChildrenData(strings.first + posStrings);
		tableEntry.strings.push_back(std::pair<std::wstring, std::wstring>(stringEntry->szKey, std::wstring(reinterpret_cast<const WCHAR* const>(stringData.first), stringEntry->Header.wValueLength)));

		posStrings += Round(stringEntry->Header.wLength);
	}

	return tableEntry;
}

OffsetLengthPair CUpdateVersionInfo::GetChildrenData(const BYTE* entryData) {
	auto entry = reinterpret_cast<const VS_VERSION_STRING*>(entryData);
	auto headerOffset = entryData;
	auto headerSize = sizeof(VS_VERSION_HEADER);
	auto keySize = (wcslen(entry->szKey) + 1) * sizeof(WCHAR);
	auto childrenOffset = Round(headerSize + keySize);

	auto pChildren = headerOffset + childrenOffset;
	auto childrenSize = entry->Header.wLength - childrenOffset;
	return OffsetLengthPair(pChildren, childrenSize);
}

BOOL CUpdateVersionInfo::ApiUpdatePEVersion(CString strFile, CString strCode, CString strValue)
{
	CString strBuffer;
	LPBYTE lpVersionData = NULL;
	TRANSLATE *lpTranslate;
	DWORD dwDataSize = 0;
	DWORD dwHandle = 0;
	try
	{
		dwDataSize = GetFileVersionInfoSize(strFile.AllocSysString(), &dwHandle);
		if (dwDataSize == 0)
		{
			throw CString("GetFileVersionInfoSize");
		}

		lpVersionData = new BYTE[dwDataSize];
		memset(lpVersionData, 0, dwDataSize);
		if (!::GetFileVersionInfo(strFile.AllocSysString(), dwHandle, dwDataSize, lpVersionData))
		{
			throw CString("GetFileVersionInfo");
		}

		VS_VERSIONINFO* pVerInfo = (VS_VERSIONINFO*)lpVersionData;
		LPBYTE pOffsetBytes = (BYTE*)(&pVerInfo->szKey[wcslen(pVerInfo->szKey)] + 1);
		VS_FIXEDFILEINFO*pFixedInfo = (VS_FIXEDFILEINFO*)roundpos(pVerInfo, pOffsetBytes, 4);

		HANDLE hResource = BeginUpdateResource(strFile.AllocSysString(), FALSE);
		if (NULL == hResource)
		{
			throw CString("BeginUpdateResource");
		}

		UINT nQuerySize;
		if (::VerQueryValue(lpVersionData, L"\\VarFileInfo\\Translation", (LPVOID*)&lpTranslate, &nQuerySize))
		{
			lpTranslate->wLanguage = 2052;//失败了则使用默认的中文解决方案
		}

		WORD ver[4] = { 0 };// 分别对应0.0.0.0
		CStringArray sArray;
		if ("FileVersion" == strCode)
		{
			this->SplitString(&sArray, strValue, L".");
			for (size_t i = 0; i < sArray.GetSize(); i++)
			{
				CStringA s = CStringTochar(sArray[i]);
				ver[i] = atoi(s);
			}

			pFixedInfo->dwFileVersionMS = MAKELONG(ver[1], ver[0]);
			pFixedInfo->dwFileVersionLS = MAKELONG(ver[3], ver[2]);
		}

		if ("ProductVersion" == strCode)
		{
			this->SplitString(&sArray, strValue, L".");
			for (size_t i = 0; i < sArray.GetSize(); i++)
			{
				CStringA s = CStringTochar(sArray[i]);
				ver[i] = atoi(s);
			}

			pFixedInfo->dwProductVersionMS = MAKELONG(ver[1], ver[0]);
			pFixedInfo->dwProductVersionLS = MAKELONG(ver[3], ver[2]);
		}

		/// 修改字符串形式的版本信息
		LPTSTR lpStringBuf = NULL;
		DWORD dwStringLen = 0;
		TCHAR szTemp[MAX_PATH] = { 0 };
		TCHAR szVersion[MAX_PATH] = { 0 };

		wsprintf(szTemp, L"\\StringFileInfo\\%04x%04x\\%s", lpTranslate->wLanguage, lpTranslate->wCodePage, strCode.AllocSysString());
		wsprintf(szVersion, L"%d\.%d\.%d\.%d", ver[0], ver[1], ver[2], ver[3]);

		if (FALSE != ::VerQueryValue(lpVersionData, szTemp, (LPVOID*)&lpStringBuf, (PUINT)&dwStringLen))
		{
			memcpy(lpStringBuf, szVersion, (_tcslen(szVersion) + 1) * sizeof(TCHAR));
		}

		if (!::UpdateResource(hResource, RT_VERSION, MAKEINTRESOURCE(VS_VERSION_INFO), lpTranslate->wLanguage, (LPVOID)lpVersionData, dwDataSize))
		{
			throw CString("UpdateResource");
		}

		if (!EndUpdateResource(hResource, FALSE))
		{
			throw CString("EndUpdateResource");
		}

		if (lpVersionData)
			delete[] lpVersionData;

		return TRUE;
	}
	catch (...)
	{
		CString strBuffer;
		strBuffer.Format(_T("ERROR %d"), GetLastError());
		MessageBox(NULL, strBuffer, NULL, MB_OK);
		if (lpVersionData)
			delete[] lpVersionData;

		return FALSE;
	}
}

CStringA CUpdateVersionInfo::CStringTochar(CString str1)
{
	// 定义一个缓冲区来保存转换后的字符串
	CStringA strA;
	// 思考一下为什么长度要 *2
	DWORD ansiLength = str1.GetLength() * 2;
	// 把 Unicode 字符串转换为 ANSI 字符串，存入 StrA 中
	WideCharToMultiByte(CP_ACP, 0, (LPCWSTR)str1, -1, strA.GetBuffer(ansiLength), ansiLength, NULL, NULL);
	// GetBuffer 必须 Release，否则后续操作无法进行
	strA.ReleaseBuffer();
	return strA;
}

BOOL CUpdateVersionInfo::SplitString(CStringArray *pArray, CString strMultiValue, CString strSymbol)
{
	pArray->RemoveAll();

	CString strBuffer, strSingle;
	strBuffer = strMultiValue;

	if (!strSymbol.IsEmpty())
	{
		CString strValue;
		strBuffer.Replace(strSymbol, L"\\\\");
		strValue.Format(L"\\%s\\", strBuffer);
		strBuffer = strValue;
	}

	LONG nStart = 0, nClose = 0;
	while (TRUE)
	{
		nStart = strBuffer.Find('\\', nClose);
		if (nStart == -1)
			break;

		nClose = strBuffer.Find('\\', nStart + 1);
		if (nClose == -1)
			break;

		strSingle = strBuffer.Mid(nStart + 1, nClose - nStart - 1);

		if (!strSingle.IsEmpty())
			pArray->Add(strSingle);
	}

	if (pArray->GetSize() == 0 && strMultiValue.GetLength() > 0)
	{
		strSingle = strMultiValue;
		pArray->Add(strSingle);
	}

	return TRUE;
}

BOOL CUpdateVersionInfo::GetApiPEInfo(CString strFile, CString strCode, CString &strValue)
{
	LPBYTE lpVersionData = NULL;
	TRANSLATE *lpTranslate;
	DWORD dwDataSize = 0;
	DWORD dwHandle = 0;
	try
	{
		dwDataSize = GetFileVersionInfoSize(strFile.AllocSysString(), &dwHandle);
		if (dwDataSize == 0)
		{
			throw CString("GetFileVersionInfoSize");
		}

		lpVersionData = new BYTE[dwDataSize];
		memset(lpVersionData, 0, dwDataSize);
		if (!::GetFileVersionInfo(strFile.AllocSysString(), dwHandle, dwDataSize, lpVersionData))
		{
			throw CString("GetFileVersionInfo");
		}

		VS_VERSIONINFO* pVerInfo = (VS_VERSIONINFO*)lpVersionData;
		LPBYTE pOffsetBytes = (BYTE*)(&pVerInfo->szKey[wcslen(pVerInfo->szKey)] + 1);
		VS_FIXEDFILEINFO*pFixedInfo = (VS_FIXEDFILEINFO*)roundpos(pVerInfo, pOffsetBytes, 4);

		UINT nQuerySize;
		if (::VerQueryValue(lpVersionData, L"\\VarFileInfo\\Translation", (LPVOID*)&lpTranslate, &nQuerySize) == FALSE)
		{
			lpTranslate->wLanguage = 2052;//失败了则使用默认的中文解决方案
		}

		LPTSTR lpStringBuf = NULL;
		DWORD dwStringLen = 0;
		TCHAR szTemp[MAX_PATH] = { 0 };

		wsprintf(szTemp, L"\\StringFileInfo\\%04x%04x\\%s", lpTranslate->wLanguage, lpTranslate->wCodePage, strCode.AllocSysString());

		if (FALSE != ::VerQueryValue(lpVersionData, szTemp, (LPVOID*)&lpStringBuf, (PUINT)&dwStringLen))
		{
			strValue.Format(_T("%s"), lpStringBuf);
		}

		if (lpVersionData)
			delete[] lpVersionData;

		return TRUE;
	}
	catch (...)
	{
		CString strBuffer;
		strBuffer.Format(_T("ERROR %d"), GetLastError());
		MessageBox(NULL, strBuffer, NULL, MB_OK);
		if (lpVersionData)
			delete[] lpVersionData;
		return FALSE;
	}
}

STDMETHODIMP CUpdateVersionInfo::UpdatePEVersion(BSTR lpFile, BSTR lpKey, BSTR lpValue, LONG* lppResult)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// TODO: 在此添加实现代码
	CString strFile, strCode, strValue;

	strFile = lpFile;
	strCode = lpKey;
	strValue = lpValue;
	BOOL rc = this->UpdatePEVersion(strFile, strCode, strValue);
	if (rc == FALSE)
	{
		*lppResult = FALSE;
		return S_FALSE;
	}
	else
	{
		*lppResult = TRUE;
	}

	return S_OK;
}
