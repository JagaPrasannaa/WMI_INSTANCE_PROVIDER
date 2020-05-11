
#ifndef _Sample_H_
#define _Sample_H_

#include <wbemprov.h>
#include <WbemIdl.h>
#pragma comment(lib, "wbemuuid.lib")

typedef LPVOID* PPVOID;
class CWbemProvider : public IWbemProviderInit, public IWbemServices
{
public:
	STDMETHODIMP QueryInterface(REFIID riid, void** ppvObject);
	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();

	STDMETHODIMP Initialize(LPWSTR wszUser, LONG lFlags, LPWSTR wszNamespace, LPWSTR wszLocale, IWbemServices* pNamespace, IWbemContext* pCtx, IWbemProviderInitSink* pInitSink);

	STDMETHODIMP OpenNamespace(const BSTR strNamespace, long lFlags, IWbemContext* pCtx, IWbemServices** ppWorkingNamespace, IWbemCallResult** ppResult);
	STDMETHODIMP CancelAsyncCall(IWbemObjectSink* pSink);
	STDMETHODIMP QueryObjectSink(long lFlags, IWbemObjectSink** ppResponseHandler);
	STDMETHODIMP GetObject(const BSTR strObjectPath, long lFlags, IWbemContext* pCtx, IWbemClassObject** ppObject, IWbemCallResult** ppCallResult);
	STDMETHODIMP GetObjectAsync(const BSTR strObjectPath, long lFlags, IWbemContext* pCtx, IWbemObjectSink* pResponseHandler);
	STDMETHODIMP PutClass(IWbemClassObject* pObject, long lFlags, IWbemContext* pCtx, IWbemCallResult** ppCallResult);
	STDMETHODIMP PutClassAsync(IWbemClassObject* pObject, long lFlags, IWbemContext* pCtx, IWbemObjectSink* pResponseHandler);
	STDMETHODIMP DeleteClass(const BSTR strClass, long lFlags, IWbemContext* pCtx, IWbemCallResult** ppCallResult);
	STDMETHODIMP DeleteClassAsync(const BSTR strClass, long lFlags, IWbemContext* pCtx, IWbemObjectSink* pResponseHandler);
	STDMETHODIMP CreateClassEnum(const BSTR strSuperclass, long lFlags, IWbemContext* pCtx, IEnumWbemClassObject** ppEnum);
	STDMETHODIMP CreateClassEnumAsync(const BSTR strSuperclass, long lFlags, IWbemContext* pCtx, IWbemObjectSink* pResponseHandler);
	STDMETHODIMP PutInstance(IWbemClassObject* pInst, long lFlags, IWbemContext* pCtx, IWbemCallResult** ppCallResult);
	STDMETHODIMP PutInstanceAsync(IWbemClassObject* pInst, long lFlags, IWbemContext* pCtx, IWbemObjectSink* pResponseHandler);
	STDMETHODIMP DeleteInstance(const BSTR strObjectPath, long lFlags, IWbemContext* pCtx, IWbemCallResult** ppCallResult);
	STDMETHODIMP DeleteInstanceAsync(const BSTR strObjectPath, long lFlags, IWbemContext* pCtx, IWbemObjectSink* pResponseHandler);
	STDMETHODIMP CreateInstanceEnum(const BSTR strFilter, long lFlags, IWbemContext* pCtx, IEnumWbemClassObject** ppEnum);
	STDMETHODIMP CreateInstanceEnumAsync(const BSTR strFilter, long lFlags, IWbemContext* pCtx, IWbemObjectSink* pResponseHandler);
	STDMETHODIMP ExecQuery(const BSTR strQueryLanguage, const BSTR strQuery, long lFlags, IWbemContext* pCtx, IEnumWbemClassObject** ppEnum);
	STDMETHODIMP ExecQueryAsync(const BSTR strQueryLanguage, const BSTR strQuery, long lFlags, IWbemContext* pCtx, IWbemObjectSink* pResponseHandler);
	STDMETHODIMP ExecNotificationQuery(const BSTR strQueryLanguage, const BSTR strQuery, long lFlags, IWbemContext* pCtx, IEnumWbemClassObject** ppEnum);
	STDMETHODIMP ExecNotificationQueryAsync(const BSTR strQueryLanguage, const BSTR strQuery, long lFlags, IWbemContext* pCtx, IWbemObjectSink* pResponseHandler);
	STDMETHODIMP ExecMethod(const BSTR strObjectPath, const BSTR strMethodName, long lFlags, IWbemContext* pCtx, IWbemClassObject* pInParams, IWbemClassObject** ppOutParams, IWbemCallResult** ppCallResult);
	STDMETHODIMP ExecMethodAsync(const BSTR strObjectPath, const BSTR strMethodName, long lFlags, IWbemContext* pCtx, IWbemClassObject* pInParams, IWbemObjectSink* pResponseHandler);

	CWbemProvider();
	~CWbemProvider();
	void SampleMethod(IWbemClassObject* pInParams, IWbemClassObject* pOutParams);

private:
	LONG          m_cRef;
	IWbemServices* m_pNamespace;
};

class CClassFactory : public IClassFactory
{
public:
	STDMETHODIMP QueryInterface(REFIID riid, void** ppvObject);
	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();

	STDMETHODIMP CreateInstance(IUnknown* pUnkOuter, REFIID riid, void** ppvObject);
	STDMETHODIMP LockServer(BOOL fLock);

private:
	LONG          m_cRef;
};


#endif