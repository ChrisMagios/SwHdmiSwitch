#define _WIN32_DCOM
#include <iostream>
using namespace std;
#include <wbemidl.h>
#include <wbemcli.h>
#include <oleauto.h>
#pragma comment(lib, "wbemuuid.lib")

const BSTR queryMouse = L"select * from win32_PointingDevice";


int main() {
    HRESULT hr;
    IWbemLocator *pWbemLoc;
    IWbemServices *pWbemServices;
    IWbemClassObject *pWbemObject = 0;
    IEnumWbemClassObject *pEnum;
    VARIANT pVal;
    CIMTYPE pType;




    if(SUCCEEDED(CoInitializeEx(NULL, COINIT_MULTITHREADED))) {
        if(SUCCEEDED(CoInitializeSecurity(NULL, -1, NULL, NULL, RPC_C_AUTHN_LEVEL_DEFAULT, RPC_C_IMP_LEVEL_IMPERSONATE, NULL, EOAC_NONE, NULL))) {
            if(SUCCEEDED(CoCreateInstance(CLSID_WbemLocator, NULL, CLSCTX_INPROC_SERVER, IID_IWbemLocator, (LPVOID *) &pWbemLoc))) {
                if(SUCCEEDED(pWbemLoc->ConnectServer(BSTR(L"\\\\DESKTOP-UQG6F0A\\root\\cimv2"), NULL, NULL, 0, NULL, 0, 0, &pWbemServices))) {
                    if(SUCCEEDED(CoSetProxyBlanket(pWbemServices, RPC_C_AUTHN_WINNT, RPC_C_AUTHZ_NONE, NULL, RPC_C_AUTHN_LEVEL_CALL, RPC_C_IMP_LEVEL_IMPERSONATE, NULL, EOAC_NONE))) {
                        // Try to get all pointing devices e.g. Mouses.
                        if(SUCCEEDED(pWbemServices->ExecQuery(BSTR(L"WQL"), queryMouse, WBEM_FLAG_DIRECT_READ, NULL, &pEnum))) {
                            //Try to get current connected mouse out of list
                            ULONG retObjects = 0;
                            if(SUCCEEDED(pEnum->Next(5000, 1, &pWbemObject, &retObjects))){
                                // Get unique mouse id
                                if(SUCCEEDED(pWbemObject->Get(BSTR(L"DeviceID"), 0, &pVal, &pType, 0))) {  
                                    wprintf(L"The DeviceID of pointing device is: %s\n.", V_BSTR(&pVal));
                                }
                                else {
                                    cout << "Couldnt retrive DeviceId for pointing device: 0x"  << endl;
                                }
                                
                            }
                            else {
                                cout << "Couldnt retrive Mouse-Object: 0x"  << endl;
                            }
                        }
                        else {
                            cout << "Couldnt retrive Point Device Object List: 0x"  << endl;
                        }
                    }
                    else {
                        pWbemLoc->Release();
                        cout << "Failed to set security levels on WMI-Connection: 0x"  << endl;
                    }
                }
                else {
                    pWbemLoc->Release();
                    cout << "Connecting to WIM-Server failed: 0x"  << endl;
                }
            }
            else {
                cout << "Creating WbemLocator failed: 0x"  << endl;
            }
        }
        else {
            cout << "COM Initialize security failed failed: 0x"  << endl;
        }
    }
    else {
        cout << "COM Initialize failed: 0x"  << endl;
    }

    pWbemLoc->Release();
    pWbemServices->Release();
     
    CoUninitialize();


    return 0;

}