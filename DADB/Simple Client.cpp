//============================================================================
// TITLE: Simple Client.cpp
//
// CONTENTS:
// 
// A client program that illustrates how to insert data to database quickly with OLE DB. 
//
// (c) Copyright 2023 websocket4opc@gmail.com
// ALL RIGHTS RESERVED.
//
// DISCLAIMER:
//  This code is provided by websocket4opc@gmail.com solely to assist in 
//  understanding and use of the lowest level DB programming in OPC classic.
//  This code is provided as-is and without warranty or support of any sort.
//
// MODIFICATION LOG:
//
// Date       By						Notes
// ---------- -----------------------	-----
// 2023/12/15 websocket4opc@gmail.com   Initial implementation.

#include "DA.h"

HRESULT listServers(CLSID& cidOpcServer)
{
	ULONG fetched = 0;
	HRESULT hr = S_OK;
	CComHeapPtr<OLECHAR> bsProgID, lpszUserType, lpszVerIndProgID;

	CATID arrcatid[3] = { NULL };
	arrcatid[0] = __uuidof(CATID_OPCDAServer10);
	arrcatid[1] = __uuidof(CATID_OPCDAServer20);
	arrcatid[2] = __uuidof(CATID_OPCDAServer30);

	CComPtr<IOPCServerList2> spIOPCServerList2;

	if (FAILED(hr = spIOPCServerList2.CoCreateInstance(__uuidof(OpcServerList), spIOPCServerList2, CLSCTX_ALL)))
	{
		printf("CoCreateInstance() for IOPCServerList2 failed\n");
		return hr;
	}

	CComPtr<IOPCEnumGUID> spEnum;

	hr = spIOPCServerList2->EnumClassesOfCategories(sizeof arrcatid / sizeof CATID, arrcatid, 0, NULL, &spEnum);

	if (spEnum.p)
	{
		while ((hr = spEnum->Next(1, &cidOpcServer, &fetched)) == S_OK)
		{
			hr = spIOPCServerList2->GetClassDetails(cidOpcServer, &bsProgID, &lpszUserType, &lpszVerIndProgID);

			if (FAILED(hr)) {
				_tprintf(_T("GetClassDetails() failed\n"));
				return hr;
			}

			break;
		}
	}

	return hr;
}

void displayResult(CSession &session) {

	CCommand<CManualAccessor> command;

	const USHORT uColumns = 4;
	CComVariant vValues[uColumns]{};

	HRESULT hr = command.CreateAccessor(uColumns, vValues, sizeof vValues);

	if (FAILED(hr))
	{
		printf("CreateAccessor() failed\n");
		return;
	}

	for (ULONG l = 0; l < uColumns; l++)
	{
		command.AddBindEntry(l + 1, DBTYPE_VARIANT, NULL, &vValues[l], NULL, NULL);
	}

	hr = command.Open(session, "select * from OPCDA", NULL, NULL);

	if (FAILED(hr))
	{
		printf("command.Open() failed\n");
		return;
	}

	ULONG count = 0;

	while (command.MoveNext() == S_OK) {

		CComVariant* pBind = (CComVariant*)command.m_pBuffer;
		
		count++;

		COleDateTime dateTime(pBind[2].date);
		printf("%S (%f, %s, %s)\n", pBind[0].bstrVal, pBind[1].fltVal, dateTime.Format("%F %T").GetString(), pBind[3].iVal == OPC_QUALITY_GOOD ? "good" : "bad");;
	}

	printf("\nTotal rows: %d\n", count);
}

int main(int argc, CHAR* argv[]) {
	
	CoInitializeEx(NULL, COINIT_MULTITHREADED);

	{
		CDataSource dataSource;
		CSession session;

		const WCHAR szUDLFile[] = L"OPCDA.udl";

		HRESULT hr = dataSource.OpenFromFileName(szUDLFile);

		if (FAILED(hr)) {

			printf("OpenFromFileName() failed\n");
			goto END;
		}

		hr = session.Open(dataSource);

		if (FAILED(hr))
		{
			printf("Open() failed\n");
			dataSource.Close();
			goto END;
		}

		CLSID cidOpcServer;

		if (FAILED(listServers(cidOpcServer)))
		{
			printf("listServers() failed\n");
			dataSource.Close();
			goto END;
		}

		if (FAILED(DA(cidOpcServer, session))) {
			printf("DA() failed\n");
			dataSource.Close();
			goto END;
		}

		printf("\nretrieving rows from database...\n\n");

		displayResult(session);

		dataSource.Close();
	}

	system("pause");

END:
	CoUninitialize();
	return(EXIT_SUCCESS);
}

