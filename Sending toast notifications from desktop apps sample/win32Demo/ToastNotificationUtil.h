#pragma once
#include "stdafx.h"

using namespace Microsoft::WRL;
using namespace ABI::Windows::UI::Notifications;
using namespace ABI::Windows::Data::Xml::Dom;
using namespace Windows::Foundation;

class ToastNotificationUtil {
	static HRESULT InstallShortcut(wchar_t *shortcutPath);
	HRESULT CreateToastXml(IToastNotificationManagerStatics *toastManager, IXmlDocument** inputXml);
	HRESULT CreateToast(IToastNotificationManagerStatics *toastManager, IXmlDocument *xml);
	HRESULT ToastNotificationUtil::SetImageSrc(_In_z_ wchar_t *imagePath, _In_ IXmlDocument *toastXml);
	HRESULT ToastNotificationUtil::SetTextValues(_In_reads_(textValuesCount) wchar_t **textValues, _In_ UINT32 textValuesCount, _In_reads_(textValuesCount) UINT32 *textValuesLengths, _In_ IXmlDocument *toastXml);
	HRESULT ToastNotificationUtil::SetNodeValueString(_In_ HSTRING inputString, _In_ IXmlNode *node, _In_ IXmlDocument *xml);
	HWND _hwnd;
	HWND _hEdit;
public:
	ToastNotificationUtil(HWND appWnd, HWND hEdit):_hwnd(appWnd), _hEdit(hEdit){}
	~ToastNotificationUtil() {}
	static HRESULT TryCreateShortcut(void);
	HRESULT DisplayToast(void);
};