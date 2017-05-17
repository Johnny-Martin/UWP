#include "stdafx.h"
#include "ToastNotificationUtil.h"
#include "StringReferenceWrapper.h"
#include "ToastEventHandler.h"

const wchar_t AppId[] = L"Microsoft.Samples.DesktopToasts";

HRESULT ToastNotificationUtil::InstallShortcut(wchar_t *shortcutPath) {
	wchar_t exePath[MAX_PATH];

	DWORD charWritten = GetModuleFileNameEx(GetCurrentProcess(), nullptr, exePath, ARRAYSIZE(exePath));

	HRESULT hr = charWritten > 0 ? S_OK : E_FAIL;

	if (SUCCEEDED(hr))
	{
		ComPtr<IShellLink> shellLink;
		hr = CoCreateInstance(CLSID_ShellLink, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&shellLink));

		if (SUCCEEDED(hr))
		{
			hr = shellLink->SetPath(exePath);
			if (SUCCEEDED(hr))
			{
				hr = shellLink->SetArguments(L"");
				if (SUCCEEDED(hr))
				{
					ComPtr<IPropertyStore> propertyStore;

					hr = shellLink.As(&propertyStore);
					if (SUCCEEDED(hr))
					{
						PROPVARIANT appIdPropVar;
						hr = InitPropVariantFromString(AppId, &appIdPropVar);
						if (SUCCEEDED(hr))
						{
							hr = propertyStore->SetValue(PKEY_AppUserModel_ID, appIdPropVar);
							if (SUCCEEDED(hr))
							{
								hr = propertyStore->Commit();
								if (SUCCEEDED(hr))
								{
									ComPtr<IPersistFile> persistFile;
									hr = shellLink.As(&persistFile);
									if (SUCCEEDED(hr))
									{
										hr = persistFile->Save(shortcutPath, TRUE);
									}
								}
							}
							PropVariantClear(&appIdPropVar);
						}
					}
				}
			}
		}
	}
	return hr;
}

HRESULT ToastNotificationUtil::TryCreateShortcut(void) {
	wchar_t shortcutPath[MAX_PATH];
	DWORD charWritten = GetEnvironmentVariable(L"APPDATA", shortcutPath, MAX_PATH);
	HRESULT hr = charWritten > 0 ? S_OK : E_INVALIDARG;

	if (SUCCEEDED(hr))
	{
		errno_t concatError = wcscat_s(shortcutPath, ARRAYSIZE(shortcutPath), L"\\Microsoft\\Windows\\Start Menu\\Programs\\Win32Demo.lnk");
		hr = concatError == 0 ? S_OK : E_INVALIDARG;
		if (SUCCEEDED(hr))
		{
			DWORD attributes = GetFileAttributes(shortcutPath);
			bool fileExists = attributes < 0xFFFFFFF;

			if (!fileExists)
			{
				hr = InstallShortcut(shortcutPath);
			}
			else
			{
				hr = S_FALSE;
			}
		}
	}
	return hr;
}

HRESULT ToastNotificationUtil::DisplayToast(void) {
	ComPtr<IToastNotificationManagerStatics> toastStatics;
	HRESULT hr = GetActivationFactory(StringReferenceWrapper(RuntimeClass_Windows_UI_Notifications_ToastNotificationManager).Get(), &toastStatics);

	if (SUCCEEDED(hr))
	{
		ComPtr<IXmlDocument> toastXml;
		hr = CreateToastXml(toastStatics.Get(), &toastXml);
		if (SUCCEEDED(hr))
		{
			hr = CreateToast(toastStatics.Get(), toastXml.Get());
		}
	}
	return hr;
}

HRESULT ToastNotificationUtil::CreateToastXml(IToastNotificationManagerStatics *toastManager, IXmlDocument** inputXml) {
	// Retrieve the template XML
	HRESULT hr = toastManager->GetTemplateContent(ToastTemplateType_ToastImageAndText04, inputXml);
	if (SUCCEEDED(hr))
	{
		wchar_t *imagePath = _wfullpath(nullptr, L"toastImageAndText.png", MAX_PATH);

		hr = imagePath != nullptr ? S_OK : HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
		if (SUCCEEDED(hr))
		{
			hr = SetImageSrc(imagePath, *inputXml);
			if (SUCCEEDED(hr))
			{
				wchar_t* textValues[] = {
					L"–¬∆¨…œœﬂ¿≤",
					L"ª∂¿÷ÀÃ2"
				};

				UINT32 textLengths[2];
				textLengths[0] = 5;
				textLengths[1] = 4;
				hr = SetTextValues(textValues, 2, textLengths, *inputXml);
			}
		}
	}
	return hr;
}

HRESULT ToastNotificationUtil::CreateToast(IToastNotificationManagerStatics *toastManager, IXmlDocument *xml) {
	ComPtr<IToastNotifier> notifier;
	HRESULT hr = toastManager->CreateToastNotifierWithId(StringReferenceWrapper(AppId).Get(), &notifier);
	if (SUCCEEDED(hr))
	{
		ComPtr<IToastNotificationFactory> factory;
		hr = GetActivationFactory(StringReferenceWrapper(RuntimeClass_Windows_UI_Notifications_ToastNotification).Get(), &factory);
		if (SUCCEEDED(hr))
		{
			ComPtr<IToastNotification> toast;
			hr = factory->CreateToastNotification(xml, &toast);
			if (SUCCEEDED(hr))
			{
				// Register the event handlers
				EventRegistrationToken activatedToken, dismissedToken, failedToken;
				ComPtr<ToastEventHandler> eventHandler(new ToastEventHandler(_hwnd, _hEdit));

				hr = toast->add_Activated(eventHandler.Get(), &activatedToken);
				if (SUCCEEDED(hr))
				{
					hr = toast->add_Dismissed(eventHandler.Get(), &dismissedToken);
					if (SUCCEEDED(hr))
					{
						hr = toast->add_Failed(eventHandler.Get(), &failedToken);
						if (SUCCEEDED(hr))
						{
							hr = notifier->Show(toast.Get());
						}
					}
				}
			}
		}
	}
	return hr;
}

// Set the value of the "src" attribute of the "image" node
HRESULT ToastNotificationUtil::SetImageSrc(_In_z_ wchar_t *imagePath, _In_ IXmlDocument *toastXml)
{
	wchar_t imageSrc[MAX_PATH] = L"file:///";
	HRESULT hr = StringCchCat(imageSrc, ARRAYSIZE(imageSrc), imagePath);
	if (SUCCEEDED(hr))
	{
		ComPtr<IXmlNodeList> nodeList;
		hr = toastXml->GetElementsByTagName(StringReferenceWrapper(L"image").Get(), &nodeList);
		if (SUCCEEDED(hr))
		{
			ComPtr<IXmlNode> imageNode;
			hr = nodeList->Item(0, &imageNode);
			if (SUCCEEDED(hr))
			{
				ComPtr<IXmlNamedNodeMap> attributes;

				hr = imageNode->get_Attributes(&attributes);
				if (SUCCEEDED(hr))
				{
					ComPtr<IXmlNode> srcAttribute;

					hr = attributes->GetNamedItem(StringReferenceWrapper(L"src").Get(), &srcAttribute);
					if (SUCCEEDED(hr))
					{
						hr = SetNodeValueString(StringReferenceWrapper(imageSrc).Get(), srcAttribute.Get(), toastXml);
					}
				}
			}
		}
	}
	return hr;
}

// Set the values of each of the text nodes
HRESULT ToastNotificationUtil::SetTextValues(_In_reads_(textValuesCount) wchar_t **textValues, _In_ UINT32 textValuesCount, _In_reads_(textValuesCount) UINT32 *textValuesLengths, _In_ IXmlDocument *toastXml)
{
	HRESULT hr = textValues != nullptr && textValuesCount > 0 ? S_OK : E_INVALIDARG;
	if (SUCCEEDED(hr))
	{
		ComPtr<IXmlNodeList> nodeList;
		hr = toastXml->GetElementsByTagName(StringReferenceWrapper(L"text").Get(), &nodeList);
		if (SUCCEEDED(hr))
		{
			UINT32 nodeListLength;
			hr = nodeList->get_Length(&nodeListLength);
			if (SUCCEEDED(hr))
			{
				hr = textValuesCount <= nodeListLength ? S_OK : E_INVALIDARG;
				if (SUCCEEDED(hr))
				{
					for (UINT32 i = 0; i < textValuesCount; i++)
					{
						ComPtr<IXmlNode> textNode;
						hr = nodeList->Item(i, &textNode);
						if (SUCCEEDED(hr))
						{
							hr = SetNodeValueString(StringReferenceWrapper(textValues[i], textValuesLengths[i]).Get(), textNode.Get(), toastXml);
						}
					}
				}
			}
		}
	}
	return hr;
}

HRESULT ToastNotificationUtil::SetNodeValueString(_In_ HSTRING inputString, _In_ IXmlNode *node, _In_ IXmlDocument *xml)
{
	ComPtr<IXmlText> inputText;

	HRESULT hr = xml->CreateTextNode(inputString, &inputText);
	if (SUCCEEDED(hr))
	{
		ComPtr<IXmlNode> inputTextNode;

		hr = inputText.As(&inputTextNode);
		if (SUCCEEDED(hr))
		{
			ComPtr<IXmlNode> pAppendedChild;
			hr = node->AppendChild(inputTextNode.Get(), &pAppendedChild);
		}
	}

	return hr;
}
