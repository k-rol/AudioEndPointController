#include <stdio.h>
#include <wchar.h>
#include <tchar.h>
#include <string>
#include <iostream>
#include "windows.h"
#include "Mmdeviceapi.h"
#include "PolicyConfig.h"
#include "Propidl.h"
#include "Functiondiscoverykeys_devpkey.h"
#include <vector>

// Format default string for outputing a device entry. The following parameters will be used in the following order:
// Index, Device Friendly Name
#define DEVICE_OUTPUT_FORMAT "Audio Device %d: %ws"

typedef struct TGlobalState
{
	HRESULT hr;
	int option;
	IMMDeviceEnumerator *pEnum;
	IMMDeviceCollection *pDevices;
	LPWSTR strDefaultDeviceID;
	IMMDevice *pCurrentDevice;
	LPCWSTR pDeviceFormatStr;
	int deviceStateFilter;
} TGlobalState;

extern "C" __declspec(dllexport) BSTR GetList();
extern "C" __declspec(dllexport) int SetAudio(int deviceid);
void createDeviceEnumerator(TGlobalState* state);
void prepareDeviceEnumerator(TGlobalState* state);
void enumerateOutputDevices(TGlobalState* state);
HRESULT printDeviceInfo(IMMDevice* pDevice, int index, LPCWSTR outFormat, LPWSTR strDefaultDeviceID);
std::wstring getDeviceProperty(IPropertyStore* pStore, const PROPERTYKEY key);
HRESULT SetDefaultAudioPlaybackDevice(LPCWSTR devID);
void invalidParameterHandler(const wchar_t* expression, const wchar_t* function, const wchar_t* file, 
	unsigned int line, uintptr_t pReserved);

std::wstring interfacelist;

extern "C"
__declspec(dllexport)
BSTR GetList()
{
	TGlobalState state;

	state.option = 0; // 0 indicates list devices.
	state.strDefaultDeviceID = '\0';
	state.pDeviceFormatStr = _T(DEVICE_OUTPUT_FORMAT);
	state.deviceStateFilter = DEVICE_STATE_ACTIVE;

	state.hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
	createDeviceEnumerator(&state);

	return SysAllocString(interfacelist.c_str());
}

extern "C"
__declspec(dllexport)
int SetAudio(int deviceid)
{
	TGlobalState state;
	state.option = deviceid;
	state.deviceStateFilter = DEVICE_STATE_ACTIVE;

	state.hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
	createDeviceEnumerator(&state);
	return state.hr;
	//return 1;
}

// Create a multimedia device enumerator.
void createDeviceEnumerator(TGlobalState* state)
{
	state->pEnum = NULL;
	state->hr = CoCreateInstance(__uuidof(MMDeviceEnumerator), NULL, CLSCTX_ALL, __uuidof(IMMDeviceEnumerator),
		(void**)&state->pEnum);
	if (SUCCEEDED(state->hr))
	{
		prepareDeviceEnumerator(state);
	}
}

// Prepare the device enumerator
void prepareDeviceEnumerator(TGlobalState* state)
{
	state->hr = state->pEnum->EnumAudioEndpoints(eRender, state->deviceStateFilter, &state->pDevices);
	if SUCCEEDED(state->hr)
	{
		enumerateOutputDevices(state);
	}
	state->pEnum->Release();
}

// Enumerate the output devices
void enumerateOutputDevices(TGlobalState* state)
{
	UINT count;
	state->pDevices->GetCount(&count);

	// If option is less than 1, list devices
	if (state->option < 1) 
	{

		// Get default device
		IMMDevice* pDefaultDevice;
		state->hr = state->pEnum->GetDefaultAudioEndpoint(eRender, eMultimedia, &pDefaultDevice);
		if (SUCCEEDED(state->hr))
		{
			
			state->hr = pDefaultDevice->GetId(&state->strDefaultDeviceID);

			// Iterate all devices
			for (int i = 1; i <= (int)count; i++)
			{
				state->hr = state->pDevices->Item(i - 1, &state->pCurrentDevice);
				if (SUCCEEDED(state->hr))
				{
					state->hr = printDeviceInfo(state->pCurrentDevice, i, state->pDeviceFormatStr,
						state->strDefaultDeviceID);
					state->pCurrentDevice->Release();
				}
				if(i < (int)count) interfacelist.append(L"\r\n");
			}
		}
	}
	// If option corresponds with the index of an audio device, set it to default
	else if (state->option <= (int)count)
	{
		state->hr = state->pDevices->Item(state->option - 1, &state->pCurrentDevice);
		if (SUCCEEDED(state->hr))
		{
			LPWSTR strID = NULL;
			state->hr = state->pCurrentDevice->GetId(&strID);
			if (SUCCEEDED(state->hr))
			{
				state->hr = SetDefaultAudioPlaybackDevice(strID);
			}
			state->pCurrentDevice->Release();
		}
	}
	// Otherwise inform user than option doesn't correspond with a device
	else
	{
		wprintf_s(_T("Error: No audio end-point device with the index '%d'.\n"), state->option);
	}
	
	state->pDevices->Release();
}

HRESULT printDeviceInfo(IMMDevice* pDevice, int index, LPCWSTR outFormat, LPWSTR strDefaultDeviceID)
{
	// Device ID
	LPWSTR strID = NULL;
	HRESULT hr = pDevice->GetId(&strID);
	if (!SUCCEEDED(hr))
	{
		return hr;
	}

	int deviceDefault = (strDefaultDeviceID != '\0' && (wcscmp(strDefaultDeviceID, strID) == 0));

	// Device state
	DWORD dwState;
	hr = pDevice->GetState(&dwState);
	if (!SUCCEEDED(hr))
	{
		return hr;
	}
		
	IPropertyStore *pStore;
	hr = pDevice->OpenPropertyStore(STGM_READ, &pStore);
	if (SUCCEEDED(hr))
	{
		std::wstring friendlyName = getDeviceProperty(pStore, PKEY_Device_FriendlyName);
		std::wstring Desc = getDeviceProperty(pStore, PKEY_Device_DeviceDesc);
		std::wstring interfaceFriendlyName = getDeviceProperty(pStore, PKEY_DeviceInterface_FriendlyName);
		
		if (SUCCEEDED(hr))
		{
			std::wstring defaultindicator;
			if (deviceDefault == 1) defaultindicator = L"*";
			interfacelist.append(defaultindicator + std::to_wstring(index) + L" : " + Desc + L" : " + interfaceFriendlyName);
		}

		pStore->Release();
	}
	return hr;
}

std::wstring getDeviceProperty(IPropertyStore* pStore, const PROPERTYKEY key)
{
	PROPVARIANT prop;
	PropVariantInit(&prop);
	HRESULT hr = pStore->GetValue(key, &prop);
	if (SUCCEEDED(hr))
	{
		std::wstring result (prop.pwszVal);
		PropVariantClear(&prop);
		return result;
	}
	else
	{
		return std::wstring (L"");
	}
}

HRESULT SetDefaultAudioPlaybackDevice(LPCWSTR devID)
{	
	IPolicyConfigVista *pPolicyConfig;
	ERole reserved = eConsole;

    HRESULT hr = CoCreateInstance(__uuidof(CPolicyConfigVistaClient), 
		NULL, CLSCTX_ALL, __uuidof(IPolicyConfigVista), (LPVOID *)&pPolicyConfig);
	if (SUCCEEDED(hr))
	{
		hr = pPolicyConfig->SetDefaultEndpoint(devID, reserved);
		pPolicyConfig->Release();
	}
	return hr;
}

void invalidParameterHandler(const wchar_t* expression,
   const wchar_t* function, 
   const wchar_t* file, 
   unsigned int line, 
   uintptr_t pReserved)
{
   wprintf_s(_T("\nError: Invalid format_str.\n"));
   exit(1);
}