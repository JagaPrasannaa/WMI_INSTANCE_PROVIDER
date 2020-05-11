#include "pch.h"
#include <windows.h>
#include <wbemidl.h>

#include "sample.h"


struct DATA {
	int nId;
	WCHAR szData[256];
} g_data[] = {
	{1, L"apple"},
	{2, L"orange"}
};
const int g_nCount = sizeof(g_data) / sizeof(g_data[0]);
const WCHAR g_szClassName[] = L"MySampProv";


// CWbemProvider


CWbemProvider::CWbemProvider()
{
	m_cRef = 1;
	m_pNamespace = NULL;
}

CWbemProvider :: ~CWbemProvider()
{
	if (m_pNamespace != NULL)
		m_pNamespace->Release();
}

STDMETHODIMP CWbemProvider::QueryInterface(REFIID riid, void** ppvObject)
{
	*ppvObject = NULL;

	if (IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, IID_IWbemProviderInit))
		*ppvObject = static_cast<IWbemProviderInit*>(this);
	else if (IsEqualIID(riid, IID_IWbemServices))
		*ppvObject = static_cast<IWbemServices*>(this);
	else
		return E_NOINTERFACE;

	AddRef();

	return S_OK;
}

STDMETHODIMP_(ULONG) CWbemProvider::AddRef()
{
	return InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) CWbemProvider::Release()
{
	if (InterlockedDecrement(&m_cRef) == 0) {
		delete this;
		return 0;
	}

	return m_cRef;
}

STDMETHODIMP CWbemProvider::Initialize(LPWSTR wszUser, LONG lFlags, LPWSTR wszNamespace, LPWSTR wszLocale, IWbemServices* pNamespace, IWbemContext* pCtx, IWbemProviderInitSink* pInitSink)
{
	if (pNamespace != NULL) {
		m_pNamespace = pNamespace;
		m_pNamespace->AddRef();
		pInitSink->SetStatus(WBEM_S_INITIALIZED, 0);

	}
	else
		pInitSink->SetStatus(WBEM_E_FAILED, 0);

	return WBEM_S_NO_ERROR;
}

STDMETHODIMP CWbemProvider::OpenNamespace(const BSTR strNamespace, long lFlags, IWbemContext* pCtx, IWbemServices** ppWorkingNamespace, IWbemCallResult** ppResult)
{
	return WBEM_E_NOT_SUPPORTED;
}

STDMETHODIMP CWbemProvider::CancelAsyncCall(IWbemObjectSink* pSink)
{
	return WBEM_E_NOT_SUPPORTED;
}

STDMETHODIMP CWbemProvider::QueryObjectSink(long lFlags, IWbemObjectSink** ppResponseHandler)
{
	return WBEM_E_NOT_SUPPORTED;
}

STDMETHODIMP CWbemProvider::GetObject(const BSTR strObjectPath, long lFlags, IWbemContext* pCtx, IWbemClassObject** ppObject, IWbemCallResult** ppCallResult)
{
	return WBEM_E_NOT_SUPPORTED;
}

STDMETHODIMP CWbemProvider::GetObjectAsync(const BSTR strObjectPath, long lFlags, IWbemContext* pCtx, IWbemObjectSink* pResponseHandler)
{
	return WBEM_E_NOT_SUPPORTED;
}

STDMETHODIMP CWbemProvider::PutClass(IWbemClassObject* pObject, long lFlags, IWbemContext* pCtx, IWbemCallResult** ppCallResult)
{
	return WBEM_E_NOT_SUPPORTED;
}

STDMETHODIMP CWbemProvider::PutClassAsync(IWbemClassObject* pObject, long lFlags, IWbemContext* pCtx, IWbemObjectSink* pResponseHandler)
{
	return WBEM_E_NOT_SUPPORTED;
}

STDMETHODIMP CWbemProvider::DeleteClass(const BSTR strClass, long lFlags, IWbemContext* pCtx, IWbemCallResult** ppCallResult)
{
	return WBEM_E_NOT_SUPPORTED;
}

STDMETHODIMP CWbemProvider::DeleteClassAsync(const BSTR strClass, long lFlags, IWbemContext* pCtx, IWbemObjectSink* pResponseHandler)
{
	return WBEM_E_NOT_SUPPORTED;
}

STDMETHODIMP CWbemProvider::CreateClassEnum(const BSTR strSuperclass, long lFlags, IWbemContext* pCtx, IEnumWbemClassObject** ppEnum)
{
	return WBEM_E_NOT_SUPPORTED;
}

STDMETHODIMP CWbemProvider::CreateClassEnumAsync(const BSTR strSuperclass, long lFlags, IWbemContext* pCtx, IWbemObjectSink* pResponseHandler)
{
	return WBEM_E_NOT_SUPPORTED;
}

STDMETHODIMP CWbemProvider::PutInstance(IWbemClassObject* pInst, long lFlags, IWbemContext* pCtx, IWbemCallResult** ppCallResult)
{
	return WBEM_E_NOT_SUPPORTED;
}

STDMETHODIMP CWbemProvider::PutInstanceAsync(IWbemClassObject* pInst, long lFlags, IWbemContext* pCtx, IWbemObjectSink* pResponseHandler)
{
	return WBEM_E_NOT_SUPPORTED;
}

STDMETHODIMP CWbemProvider::DeleteInstance(const BSTR strObjectPath, long lFlags, IWbemContext* pCtx, IWbemCallResult** ppCallResult)
{
	return WBEM_E_NOT_SUPPORTED;
}

STDMETHODIMP CWbemProvider::DeleteInstanceAsync(const BSTR strObjectPath, long lFlags, IWbemContext* pCtx, IWbemObjectSink* pResponseHandler)
{
	return WBEM_E_NOT_SUPPORTED;
}

STDMETHODIMP CWbemProvider::CreateInstanceEnum(const BSTR strFilter, long lFlags, IWbemContext* pCtx, IEnumWbemClassObject** ppEnum)
{
	return WBEM_E_NOT_SUPPORTED;
}

STDMETHODIMP CWbemProvider::CreateInstanceEnumAsync(const BSTR strFilter, long lFlags, IWbemContext* pCtx, IWbemObjectSink* pResponseHandler)
{
	int              i;
	IWbemClassObject* pClass = NULL;
	IWbemClassObject* pNewInst[100];
	VARIANT var;

	for (i = 0; i < g_nCount; i++) {
		m_pNamespace->GetObject(strFilter, 0, pCtx, &pClass, NULL);
		pClass->SpawnInstance(0, &pNewInst[i]);
		pClass->Release();

		var.vt = VT_I4;
		var.lVal = g_data[i].nId;
		pNewInst[i]->Put(L"id", 0, &var, 0);

		var.vt = VT_BSTR;
		var.bstrVal = SysAllocString(g_data[i].szData);
		pNewInst[i]->Put(L"data", 0, &var, 0);
		VariantClear(&var);
	}

	pResponseHandler->Indicate(g_nCount, pNewInst);

	pResponseHandler->SetStatus(0, WBEM_S_NO_ERROR, NULL, NULL);

	return WBEM_S_NO_ERROR;
}

STDMETHODIMP CWbemProvider::ExecQuery(const BSTR strQueryLanguage, const BSTR strQuery, long lFlags, IWbemContext* pCtx, IEnumWbemClassObject** ppEnum)
{
	return WBEM_E_NOT_SUPPORTED;
}

STDMETHODIMP CWbemProvider::ExecQueryAsync(const BSTR strQueryLanguage, const BSTR strQuery, long lFlags, IWbemContext* pCtx, IWbemObjectSink* pResponseHandler)
{
	return WBEM_E_NOT_SUPPORTED;
}

STDMETHODIMP CWbemProvider::ExecNotificationQuery(const BSTR strQueryLanguage, const BSTR strQuery, long lFlags, IWbemContext* pCtx, IEnumWbemClassObject** ppEnum)
{
	return WBEM_E_NOT_SUPPORTED;
}

STDMETHODIMP CWbemProvider::ExecNotificationQueryAsync(const BSTR strQueryLanguage, const BSTR strQuery, long lFlags, IWbemContext* pCtx, IWbemObjectSink* pResponseHandler)
{
	return WBEM_E_NOT_SUPPORTED;
}

STDMETHODIMP CWbemProvider::ExecMethod(const BSTR strObjectPath, const BSTR strMethodName, long lFlags, IWbemContext* pCtx, IWbemClassObject* pInParams, IWbemClassObject** ppOutParams, IWbemCallResult** ppCallResult)
{
	return WBEM_E_NOT_SUPPORTED;
}

STDMETHODIMP CWbemProvider::ExecMethodAsync(const BSTR strObjectPath, const BSTR strMethodName, long lFlags, IWbemContext* pCtx, IWbemClassObject* pInParams, IWbemObjectSink* pResponseHandler)
{
	HRESULT          hr;
	IWbemClassObject* pClass = NULL;
	IWbemClassObject* pOutClass = NULL;
	IWbemClassObject* pOutParams = NULL;
	BSTR             bstrClassName;

	bstrClassName = SysAllocString(g_szClassName);
	m_pNamespace->GetObjectW(bstrClassName, 0, pCtx, &pClass, NULL);
	if (pClass == NULL) {
		SysFreeString(bstrClassName);
		pResponseHandler->SetStatus(0, WBEM_E_NOT_SUPPORTED, NULL, NULL);
		return WBEM_E_NOT_SUPPORTED;
	}

	pClass->GetMethod(strMethodName, 0, NULL, &pOutClass);
	pOutClass->SpawnInstance(0, &pOutParams);

	if (lstrcmpW(L"SampleMethod", strMethodName) == 0) {
		SampleMethod(pInParams, pOutParams);
		pResponseHandler->Indicate(1, &pOutParams);
		hr = WBEM_S_NO_ERROR;
	}
	else
		hr = WBEM_E_NOT_SUPPORTED;

	pResponseHandler->SetStatus(0, hr, NULL, NULL);

	pOutParams->Release();
	pOutClass->Release();
	pClass->Release();
	SysFreeString(bstrClassName);

	return hr;
}

void CWbemProvider::SampleMethod(IWbemClassObject* pInParams, IWbemClassObject* pOutParams)
{
	BSTR    bstrInputArgName = SysAllocString(L"strIn");
	BSTR    bstrOutputArgName = SysAllocString(L"strOut");
	BSTR    bstrRetValName = SysAllocString(L"ReturnValue");
	VARIANT varIn, varOut, varRet;
	BSTR    bstr;
	LONG    lRetValue;

	VariantInit(&varIn);
	pInParams->Get(bstrInputArgName, 0, &varIn, NULL, NULL);
	if (varIn.vt == VT_BSTR) {
		bstr = varIn.bstrVal;
		CharUpperBuffW(bstr, lstrlenW(bstr));

		varOut.vt = VT_BSTR;
		varOut.bstrVal = bstr;
		pOutParams->Put(bstrOutputArgName, 0, &varOut, 0);

		lRetValue = lstrlenW(bstr);

		VariantClear(&varIn);
		VariantClear(&varOut);
	}
	else
		lRetValue = -1;

	varRet.vt = VT_I4;
	varRet.lVal = lRetValue;
	pOutParams->Put(bstrRetValName, 0, &varRet, 0);

	SysFreeString(bstrInputArgName);
	SysFreeString(bstrOutputArgName);
	SysFreeString(bstrRetValName);
}