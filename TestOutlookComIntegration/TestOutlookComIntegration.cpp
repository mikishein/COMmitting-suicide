// TestOutlookComIntegration.cpp : This file contains the 'main' function. Program execution begins and ends there.
//
#/*import "C:\\Program Files\\Microsoft Office\\root\\vfs\\ProgramFilesCommonX64\\Microsoft Shared\\OFFICE16\\MSO.DLL" named_guids
#import "C:\\Program Files\\Microsoft Office\\root\\Office16\\MSOUTL.OLB" no_namespace  rename("GetOrganizer", "GetOrganizerAE") rename("Folder", "OlkFolder") rename("CopyFile", "OlkCopyFile")*/
#include <iostream>
#include <Windows.h>
#include <atlstr.h>
#pragma warning(disable:4996)
#include <ole2.h> // OLE2 Definitions


#define RULE_RECEIVE 0
#define FOLDER_INBOX 6
// AutoWrap() - Automation helper function...
HRESULT AutoWrap(int autoType, VARIANT* pvResult, IDispatch* pDisp, LPOLESTR ptName, int cArgs...) {
    // Begin variable-argument list...
    va_list marker;
    va_start(marker, cArgs);

    if (!pDisp) {
        MessageBoxA(NULL, "NULL IDispatch passed to AutoWrap()", "Error", 0x10010);
        _exit(0);
    }

    // Variables used...
    DISPPARAMS dp = { NULL, NULL, 0, 0 };
    DISPID dispidNamed = DISPID_PROPERTYPUT;
    DISPID dispID;
    HRESULT hr;
    char buf[200];
    char szName[200];


    // Convert down to ANSI
    WideCharToMultiByte(CP_ACP, 0, ptName, -1, szName, 256, NULL, NULL);

    // Get DISPID for name passed...
    hr = pDisp->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
    if (FAILED(hr)) {
        sprintf(buf, "IDispatch::GetIDsOfNames(\"%s\") failed w/err 0x%08lx", szName, hr);
        MessageBoxA(NULL, buf, "AutoWrap()", 0x10010);
        _exit(0);
        return hr;
    }

    // Allocate memory for arguments...
    VARIANT* pArgs = new VARIANT[cArgs + 1];
    // Extract arguments...
    for (int i = 0; i < cArgs; i++) {
        pArgs[i] = va_arg(marker, VARIANT);
    }

    // Build DISPPARAMS
    dp.cArgs = cArgs;
    dp.rgvarg = pArgs;

    // Handle special-case for property-puts!
    if (autoType & DISPATCH_PROPERTYPUT) {
        dp.cNamedArgs = 1;
        dp.rgdispidNamedArgs = &dispidNamed;
    }

    // Make the call!
    hr = pDisp->Invoke(dispID, IID_NULL, LOCALE_SYSTEM_DEFAULT, autoType, &dp, pvResult, NULL, NULL);
    if (FAILED(hr)) {
        sprintf(buf, "IDispatch::Invoke(\"%s\"=%08lx) failed w/err 0x%08lx", szName, dispID, hr);
        MessageBoxA(NULL, buf, "AutoWrap()", 0x10010);
        _exit(0);
        return hr;
    }
    // End variable-argument section...
    va_end(marker);

    delete[] pArgs;

    return hr;
}
int main()
{
    // Initialize COM for this thread...
    CoInitialize(NULL);

    // Get CLSID for our server...
    CLSID clsid;
    HRESULT hr = CLSIDFromProgID(L"Outlook.Application", &clsid);

    if (FAILED(hr)) {

        ::MessageBoxA(NULL, "CLSIDFromProgID() failed", "Error", 0x10010);
        return -1;
    }

    // Start server and get IDispatch...
    IDispatch* pXlApp;
    hr = CoCreateInstance(clsid, NULL, CLSCTX_LOCAL_SERVER, IID_IDispatch, (void**)&pXlApp);
    if (FAILED(hr)) {
        ::MessageBoxA(NULL, "Excel not registered properly", "Error", 0x10010);
        return -2;
    }

    // Get Workbooks collection
    IDispatch* pGetNameSpace;
    {
        VARIANT param;
        param.vt = VT_BSTR;
        param.bstrVal = ::SysAllocString(L"MAPI");
        VARIANT result;
        VariantInit(&result);
        
        AutoWrap(DISPATCH_PROPERTYGET, &result, pXlApp, (LPOLESTR)L"GetNamespace", 1, param);
        pGetNameSpace = result.pdispVal;
    }
    IDispatch* pGetDefaultFolder;
    {
        VARIANT x;
        x.vt = VT_I4;
        x.lVal = FOLDER_INBOX;
        VARIANT result;
        VariantInit(&result);
        AutoWrap(DISPATCH_PROPERTYGET, &result, pGetNameSpace, (LPOLESTR)L"GetDefaultFolder", 1,x);
        pGetDefaultFolder = result.pdispVal;
    }
    IDispatch* pSession;
    {
        VARIANT result;
        VariantInit(&result);
        AutoWrap(DISPATCH_PROPERTYGET, &result, pGetNameSpace, (LPOLESTR)L"Session", 0);
        pSession = result.pdispVal;
    }
    IDispatch* pDefaultStore;
    {
        VARIANT result;
        VariantInit(&result);
        AutoWrap(DISPATCH_PROPERTYGET, &result, pGetNameSpace, (LPOLESTR)L"DefaultStore", 0);
        pDefaultStore = result.pdispVal;
    }
    IDispatch* pGetRules;
    {
        VARIANT result;
        VariantInit(&result);
        AutoWrap(DISPATCH_PROPERTYGET, &result, pDefaultStore, (LPOLESTR)L"GetRules", 0);
        pGetRules = result.pdispVal;
    }
    IDispatch* pSave;
    {
        VARIANT x;
        x.vt = VT_BOOL;
        x.boolVal = false;
        VARIANT result;
        VariantInit(&result);
        AutoWrap(DISPATCH_PROPERTYGET, &result, pGetRules, (LPOLESTR)L"Save",1, x);
        pSave = result.pdispVal;
        printf("stop");
    }
    IDispatch* pRule;
    {
        VARIANT ruleName;
        ruleName.vt = VT_BSTR;
        ruleName.bstrVal = ::SysAllocString(L"COM FUCKING OBJECT");
        VARIANT ruleRecive;
        ruleRecive.vt = VT_I4;
        ruleRecive.lVal = RULE_RECEIVE;
        VARIANT result;
        VariantInit(&result);
        // For some reason parameters are pushed reversed!!!!! I'm so lucky i read it somewhere otherwise the project would have died
        AutoWrap(DISPATCH_PROPERTYGET, &result, pGetRules, (LPOLESTR)L"Create", 2,ruleRecive, ruleName);
        pRule = result.pdispVal;
    }
    IDispatch* pConditions;
    {
        VARIANT result;
        VariantInit(&result);
        AutoWrap(DISPATCH_PROPERTYGET, &result, pRule, (LPOLESTR)L"Conditions", 0);
        pConditions = result.pdispVal;
    }
    IDispatch* pSubject;
    {
        VARIANT result;
        VariantInit(&result);
        AutoWrap(DISPATCH_PROPERTYGET, &result, pConditions, (LPOLESTR)L"Subject", 0);
        pSubject = result.pdispVal;
    }
    IDispatch* pText;
    {
        VARIANT y;
        y.vt = VT_BSTR;
        y.bstrVal = ::SysAllocString(L"PullTheTrigger");
        VARIANT z;
        z.vt = VT_BSTR;
        z.bstrVal = ::SysAllocString(L"EiniYacholOd");
        VARIANT result;
        VariantInit(&result);
        AutoWrap(DISPATCH_PROPERTYPUT,NULL, pSubject, (LPOLESTR)L"Text",2,y,z);
    }
    printf("a");

   

    //// Tell Excel to quit (i.e. App.Quit)
    //AutoWrap(DISPATCH_METHOD, NULL, pXlApp, (LPOLESTR)L"Quit", 0);

    //// Release references...
    //pXlRange->Release();
    //pXlSheet->Release();
    //pXlBook->Release();
    //pXlBooks->Release();
    //pXlApp->Release();
    //VariantClear(&arr);

    // Uninitialize COM for this thread...
    CoUninitialize();

}

