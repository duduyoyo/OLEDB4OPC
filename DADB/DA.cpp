#include "DA.h"

class DataCallback : public IOPCDataCallback
{
private:
	ULONG m_ulRefs = 0;
	CSession session;
public:

	DataCallback(CSession& session)
	{
		this->session = session;
	}

	//==========================================================================
	// IUnknown

	// QueryInterface
	STDMETHODIMP QueryInterface(REFIID iid, LPVOID* ppInterface)
	{
		if (ppInterface == NULL)
		{
			return E_INVALIDARG;
		}

		if (iid == __uuidof(IUnknown))
		{
			*ppInterface = dynamic_cast<IUnknown*>(this);
			AddRef();
			return S_OK;
		}

		if (iid == __uuidof(IOPCDataCallback))
		{
			*ppInterface = dynamic_cast<IOPCDataCallback*>(this);
			AddRef();
			return S_OK;
		}

		return E_NOINTERFACE;
	}

	// AddRef
	STDMETHODIMP_(ULONG) AddRef()
	{
		return InterlockedIncrement((LONG*)& m_ulRefs);
	}

	// Release
	STDMETHODIMP_(ULONG) Release()
	{
		ULONG ulRefs = InterlockedDecrement((LONG*)& m_ulRefs);

		if (ulRefs == 0)
		{
			delete this;
		}

		return ulRefs;
	}

	//==========================================================================
	// IOPCDataCallback

	// OnDataChange
	STDMETHODIMP OnDataChange(
		DWORD       dwTransid,
		OPCHANDLE   hGroup,
		HRESULT     hrMasterquality,
		HRESULT     hrMastererror,
		DWORD       dwCount,
		OPCHANDLE* phClientItems,
		VARIANT* pvValues,
		WORD* pwQualities,
		FILETIME* pftTimeStamps,
		HRESULT* pErrors
	)
	{
		CCommand<CManualAccessor> command;


		HRESULT hr = command.CreateAccessor(0, nullptr, 0);

		if (FAILED(hr)) {
			printf("command.CreateAccessor() failed");
			return hr;
		}

		CComVariant vVariant[4];

		vVariant[0].vt = VT_BSTR;
		vVariant[1].vt = VT_R4;
		vVariant[2].vt = VT_DATE;
		vVariant[3].vt = VT_UINT;

		hr = command.CreateParameterAccessor(4, vVariant, sizeof vVariant); 

		if (FAILED(hr)) {
			printf("command.CreateParameterAccessor() failed");
			return hr;
		}

		for (DWORD ii = 0; ii < dwCount; ii++)
		{
			CComVariant vValue;
			WORD quality = pwQualities[ii] & OPC_QUALITY_MASK;

			COleDateTime oleTime = COleDateTime(pftTimeStamps[ii]);
			SYSTEMTIME st;
			oleTime.GetAsSystemTime(st);

			if (phClientItems[ii] == 0)
				CComBSTR("Random.Int1").CopyTo(&vVariant[0].bstrVal);
			if (phClientItems[ii] == 1)
				CComBSTR("Random.Int2").CopyTo(&vVariant[0].bstrVal);
			else if (phClientItems[ii] == 2)
				CComBSTR("Random.Real8").CopyTo(&vVariant[0].bstrVal);

			vVariant[1].fltVal = (FLOAT)pvValues[ii].dblVal;
			vVariant[2].date = oleTime;
			vVariant[3].iVal = quality;

			command.m_nCurrentParameter = 0;

			command.AddParameterEntry(1, DBTYPE_BSTR, NULL, &vVariant[0].bstrVal);
			command.AddParameterEntry(2, DBTYPE_R4, NULL, &vVariant[1].fltVal);
			command.AddParameterEntry(3, DBTYPE_DATE, NULL, &vVariant[2].date);
			command.AddParameterEntry(4, DBTYPE_UI2, NULL, &vVariant[3].iVal);
			
			/*
			This is not the most efficient and fastest way to insert a row to database due to query building/parsing and commit each time.
			To bulk insert, interface of IRowsetFastLoad has to be used and it is quite different from this code example.
			Contact developer to have a code example using IRowsetFastLoad, so you can completely understand the big difference between 
			IDBInitialize and IDataInitialize interfaces when trying to get a pointer to IRowsetFastLoad.
			*/

			hr = command.Open(session, "insert into OPCDA (Tag, Value, Time, Quality) Values (?,?,?,?)", NULL, NULL);

			if (FAILED(hr)) {
				printf("command.Open() failed");
				break;
			}
			else
				printf("\nOnDataChange: %S (%f, %s.%d, %s)", vVariant[0].bstrVal, vVariant[1].fltVal, oleTime.Format("%F %T").GetString(), st.wMilliseconds, quality == OPC_QUALITY_GOOD ? "good" : "bad");

			SysFreeString(vVariant[0].bstrVal);
		}
		
		return hr;
	}

	// OnReadComplete
	STDMETHODIMP OnReadComplete(
		DWORD       dwTransid,
		OPCHANDLE   hGroup,
		HRESULT     hrMasterquality,
		HRESULT     hrMastererror,
		DWORD       dwCount,
		OPCHANDLE* phClientItems,
		VARIANT* pvValues,
		WORD* pwQualities,
		FILETIME* pftTimeStamps,
		HRESULT* pErrors
	)
	{
		return E_NOTIMPL;
	}

	// OnWriteComplete
	STDMETHODIMP OnWriteComplete(
		DWORD       dwTransid,
		OPCHANDLE   hGroup,
		HRESULT     hrMastererr,
		DWORD       dwCount,
		OPCHANDLE* pClienthandles,
		HRESULT* pErrors
	)
	{
		return E_NOTIMPL;
	}


	// OnCancelComplete
	STDMETHODIMP OnCancelComplete(
		DWORD       dwTransid,
		OPCHANDLE   hGroup
	)
	{
		return E_NOTIMPL;
	}
};

HRESULT addItems(IOPCItemMgt* pOPCItemMgt) {

	OPCITEMDEF pItems[3] = { NULL };

	pItems[0].szItemID = L"random.int1";
	pItems[0].bActive = TRUE;
	pItems[0].hClient = 0;

	pItems[1].szItemID = L"random.int2";
	pItems[1].bActive = TRUE;
	pItems[1].hClient = 1;

	pItems[2].szItemID = L"random.Real8";
	pItems[2].bActive = TRUE;
	pItems[2].hClient = 2;

	OPCITEMRESULT* pResults = NULL;
	HRESULT* pErrors = NULL;

	HRESULT hResult = pOPCItemMgt->AddItems(sizeof pItems/sizeof OPCITEMDEF, pItems, &pResults, &pErrors);

	if (FAILED(hResult) || hResult == S_FALSE)
	{
		printf("AddItems() failed\n");
		CoTaskMemFree(pResults);
		CoTaskMemFree(pErrors); 
		return hResult;
	}

	CoTaskMemFree(pResults);
	CoTaskMemFree(pErrors);

	return hResult;
}

HRESULT DA(CLSID& cidOpcServer, CSession& session) {

	CComPtr<IOPCServer> pIOPCServer;

	HRESULT hr = pIOPCServer.CoCreateInstance(cidOpcServer, pIOPCServer, CLSCTX_ALL);

	if (FAILED(hr)) {
		printf("CoCreateInstance() for IOPCServer failed\n");
		return E_FAIL;
	}

	DWORD dwRevisedUpdateRate = 0;
	OPCHANDLE hGroup = 0;
	CComPtr<IOPCItemMgt> pOPCItemMgt;

	 hr = pIOPCServer->AddGroup(L"", TRUE, 1000, NULL, NULL, NULL, LOCALE_SYSTEM_DEFAULT, &hGroup, &dwRevisedUpdateRate, __uuidof(IOPCItemMgt), (LPUNKNOWN*)&pOPCItemMgt);

	 if (FAILED(hr)) {
		 printf("AddGroup() failed\n");
		 return E_FAIL;
	 }
	 
	 DataCallback* pDataCallback = new DataCallback(session);
	 pDataCallback->AddRef();
	 DWORD m_dwCookie;
	 AtlAdvise(pOPCItemMgt, pDataCallback, __uuidof(IOPCDataCallback), &m_dwCookie);

	 hr = addItems(pOPCItemMgt);

	 if (FAILED(hr)) {
		 printf("addItems() failed\n");
		 return E_FAIL;
	 }

	 printf("\npress any key to complete inserting rows to database\n");
	 getchar();

	 AtlUnadvise(pOPCItemMgt, __uuidof(IOPCDataCallback), m_dwCookie);
	 pDataCallback->Release();

	return S_OK;
}