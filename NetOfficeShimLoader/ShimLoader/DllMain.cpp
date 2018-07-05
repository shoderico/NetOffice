#include "stdafx.h"
#include "ShimProxyFactory.h"
#include "Vars.hpp"
#include "DllRegister32.h"
#include "DllRegister64.h"
#include "DllRegister32On64.h"
#include "ShimArguments.h"

HINSTANCE _module = NULL;   // DLL module handle
ULONG _components = 0;      // Count of active components
ULONG _locks = 0;			// Count of server locks

bool Is64BitWindows(bool & isWindows64bit);

BOOL APIENTRY DllMain( HMODULE hModule,
                       DWORD  ul_reason_for_call,
                       LPVOID lpReserved
                     )
{
    switch (ul_reason_for_call)
    {
		case DLL_PROCESS_ATTACH:
		{
			_module = hModule;
			::DisableThreadLibraryCalls(hModule);
			//ShimArguments* args = new ShimArguments();
			//args->Load();
			//delete args;
			break;
		}
		case DLL_THREAD_ATTACH:
		case DLL_THREAD_DETACH:
		case DLL_PROCESS_DETACH:
			break;
    }
    return TRUE;
}

STDAPI DllCanUnloadNow()
{
	HRESULT hr = (_components == 0 && _locks == 0) ? S_OK : S_FALSE;

#ifdef DEBUG

	if(S_OK != hr)
	{
		WCHAR szBuffer[128];
		wsprintf(szBuffer, L"Unexpected: %ld components left.", _components);
		MessageBox(GetDesktopWindow(), szBuffer, L"DllCanUnloadNow", 0);
	}

#endif

	return hr;
}

STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, void** ppv)
{
	*ppv = NULL;

	if (rclsid != ShimProxy_CLSID)
		return CLASS_E_CLASSNOTAVAILABLE;

	ShimProxyFactory* pCF = new (std::nothrow) ShimProxyFactory();
	if (NULL == pCF)
		return E_OUTOFMEMORY;

	HRESULT hr = pCF->QueryInterface(riid, ppv);
	if (FAILED(hr))
	{
		*ppv = NULL;
		delete pCF;
	}

	return hr;
}

STDAPI DllRegisterServer()
{
	HRESULT hr = S_OK;

	if (ENABLE_SELF_REGISTRATION)
	{
		try
		{
			#if X64BUILD

				hr = ShimLoader_Register64::DllRegister(
					_module,
					ShimProxy_Host_Application,
					ShimProxy_LoadBehavior,
					ShimProxy_CommandLineSafe,
					ShimProxy_ProgID,
					ShimProxy_CLSID_Text,
					ShimProxy_FriendlyName,
					ShimProxy_Description,
					ShimProxy_Version,
					static_cast<RegisterMode>(SELF_REGISTER_MODE));

			#else

				bool is64BitWindows = false;
				if (Is64BitWindows(is64BitWindows))
				{
					if (is64BitWindows)
					{
						hr = ShimLoader_Register32On64::DllRegister(
							_module,
							ShimProxy_Host_Application,
							ShimProxy_LoadBehavior,
							ShimProxy_CommandLineSafe,
							ShimProxy_ProgID,
							ShimProxy_CLSID_Text,
							ShimProxy_FriendlyName,
							ShimProxy_Description,
							ShimProxy_Version,
							static_cast<RegisterMode>(SELF_REGISTER_MODE));
					}
					else
					{
						hr = ShimLoader_Register32::DllRegister(
							_module,
							ShimProxy_Host_Application,
							ShimProxy_LoadBehavior,
							ShimProxy_CommandLineSafe,
							ShimProxy_ProgID,
							ShimProxy_CLSID_Text,
							ShimProxy_FriendlyName,
							ShimProxy_Description,
							ShimProxy_Version,
							static_cast<RegisterMode>(SELF_REGISTER_MODE));
					}
				}
				else
				{
					hr = E_FAIL;
				}

			#endif
		}
		catch (...)
		{
			hr = E_FAIL;
		}
	}

	return hr;
}

STDAPI DllUnregisterServer()
{
	HRESULT hr = S_OK;

	if (ENABLE_SELF_REGISTRATION)
	{
		try
		{
			#if X64BUILD

				hr = ShimLoader_Register64::DllUnregister(
					ShimProxy_Host_Application,
					ShimProxy_ProgID,
					ShimProxy_CLSID_Text,
					ShimProxy_Version,
					static_cast<RegisterMode>(SELF_REGISTER_MODE));

			#else

			bool is64BitWindows = false;
			if (Is64BitWindows(is64BitWindows))
			{
				hr = ShimLoader_Register32On64::DllUnregister(
					ShimProxy_Host_Application,
					ShimProxy_ProgID,
					ShimProxy_CLSID_Text,
					ShimProxy_Version,
					static_cast<RegisterMode>(SELF_REGISTER_MODE));
			}
			else
			{
				hr = ShimLoader_Register32::DllUnregister(
					ShimProxy_Host_Application,
					ShimProxy_ProgID,
					ShimProxy_CLSID_Text,
					ShimProxy_Version,
					static_cast<RegisterMode>(SELF_REGISTER_MODE));
			}

			#endif
		}
		catch (...)
		{
			hr = E_FAIL;
		}
	}

	return hr;
}

bool Is64BitWindows(bool & isWindows64bit)
{
#if _WIN64

	isWindows64bit = true;
	return true;

#elif _WIN32

	BOOL isWow64 = FALSE;
	LPFN_ISWOW64PROCESS fnIsWow64Process = (LPFN_ISWOW64PROCESS)GetProcAddress(GetModuleHandle(TEXT("kernel32")), "IsWow64Process");

	if (fnIsWow64Process)
	{
		if (!fnIsWow64Process(GetCurrentProcess(), &isWow64))
			return false;

		if (isWow64)
			isWindows64bit = true;
		else
			isWindows64bit = false;

		return true;
	}
	else
		return false;

#else

	assert(0);
	return false;

#endif
}
