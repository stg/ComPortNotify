// WinRT toast notifications (Windows 10/11+) using raw ABI interfaces.
// Fire-and-forget: shows a title + body.
//
// Not wired into the app yet.

#include <windows.h>
#include <objbase.h>
#include <roapi.h>
#include <winstring.h>
#include <cwchar>
#include <windows.ui.notifications.h>
#include <windows.data.xml.dom.h>

static const wchar_t *TOAST_AUMID = L"DSp.Tools.CPNotify.1";
static HRESULT g_toast_last_hr = S_OK;

extern "C" HRESULT toast_last_error() {
    return g_toast_last_hr;
}

static wchar_t *xml_escape(const wchar_t *src) {
    if(!src) {
        wchar_t *empty = (wchar_t *)malloc(sizeof(wchar_t));
        if(empty) empty[0] = L'\0';
        return empty;
    }

    size_t extra = 0;
    for(const wchar_t *p = src; *p; ++p) {
        switch(*p) {
            case L'&': extra += 4; break;   // &amp;
            case L'<': extra += 3; break;   // &lt;
            case L'>': extra += 3; break;   // &gt;
            case L'\"': extra += 5; break;  // &quot;
            case L'\'': extra += 5; break;  // &apos;
            default: break;
        }
    }

    size_t len = wcslen(src);
    wchar_t *out = (wchar_t *)malloc(sizeof(wchar_t) * (len + extra + 1));
    if(!out) return NULL;

    wchar_t *dst = out;
    for(const wchar_t *p = src; *p; ++p) {
        switch(*p) {
            case L'&': wcscpy(dst, L"&amp;"); dst += 5; break;
            case L'<': wcscpy(dst, L"&lt;"); dst += 4; break;
            case L'>': wcscpy(dst, L"&gt;"); dst += 4; break;
            case L'\"': wcscpy(dst, L"&quot;"); dst += 6; break;
            case L'\'': wcscpy(dst, L"&apos;"); dst += 6; break;
            default: *dst++ = *p; break;
        }
    }
    *dst = L'\0';
    return out;
}

static wchar_t *build_toast_xml(const wchar_t *title, const wchar_t *body) {
    wchar_t *et = xml_escape(title);
    wchar_t *eb = xml_escape(body);
    if(!et || !eb) {
        if(et) free(et);
        if(eb) free(eb);
        return NULL;
    }

    const wchar_t *fmt = L"<toast><visual><binding template=\"ToastText02\">"
                         L"<text id=\"1\">%s</text>"
                         L"<text id=\"2\">%s</text>"
                         L"</binding></visual></toast>";
    size_t needed = wcslen(fmt) + wcslen(et) + wcslen(eb) + 1;
    wchar_t *xml = (wchar_t *)malloc(sizeof(wchar_t) * needed);
    if(xml) {
        swprintf(xml, needed, fmt, et, eb);
    }
    free(et);
    free(eb);
    return xml;
}

extern "C" bool toast_show_winrt(const wchar_t *title, const wchar_t *body) {
    HRESULT hr = RoInitialize(RO_INIT_MULTITHREADED);
    bool do_uninit = SUCCEEDED(hr);
    if(FAILED(hr) && hr != RPC_E_CHANGED_MODE) return false;

    bool ok = false;
    HSTRING hManager = NULL;
    HSTRING hToastClass = NULL;
    HSTRING hXml = NULL;
    HSTRING hAumid = NULL;
    wchar_t *xml = NULL;

    ABI::Windows::UI::Notifications::IToastNotificationManagerStatics *manager = NULL;
    ABI::Windows::UI::Notifications::IToastNotificationFactory *factory = NULL;
    ABI::Windows::UI::Notifications::IToastNotifier *notifier = NULL;
    ABI::Windows::UI::Notifications::IToastNotification *toast = NULL;
    ABI::Windows::Data::Xml::Dom::IXmlDocument *xmlDoc = NULL;
    ABI::Windows::Data::Xml::Dom::IXmlDocumentIO *xmlIO = NULL;

    hr = WindowsCreateString(L"Windows.UI.Notifications.ToastNotificationManager",
                             (UINT32)wcslen(L"Windows.UI.Notifications.ToastNotificationManager"),
                             &hManager);
    if(FAILED(hr)) goto cleanup;

    hr = RoGetActivationFactory(hManager, __uuidof(ABI::Windows::UI::Notifications::IToastNotificationManagerStatics), (void **)&manager);
    if(FAILED(hr)) goto cleanup;

    hr = manager->GetTemplateContent(
        ABI::Windows::UI::Notifications::ToastTemplateType::ToastTemplateType_ToastText02,
        &xmlDoc);
    if(FAILED(hr)) goto cleanup;

    hr = xmlDoc->QueryInterface(__uuidof(ABI::Windows::Data::Xml::Dom::IXmlDocumentIO), (void **)&xmlIO);
    if(FAILED(hr)) goto cleanup;

    xml = build_toast_xml(title ? title : L"", body ? body : L"");
    if(!xml) goto cleanup;

    hr = WindowsCreateString(xml, (UINT32)wcslen(xml), &hXml);
    if(FAILED(hr)) goto cleanup;

    hr = xmlIO->LoadXml(hXml);
    if(FAILED(hr)) goto cleanup;

    hr = WindowsCreateString(L"Windows.UI.Notifications.ToastNotification",
                             (UINT32)wcslen(L"Windows.UI.Notifications.ToastNotification"),
                             &hToastClass);
    if(FAILED(hr)) goto cleanup;

    hr = RoGetActivationFactory(hToastClass, __uuidof(ABI::Windows::UI::Notifications::IToastNotificationFactory), (void **)&factory);
    if(FAILED(hr)) goto cleanup;

    hr = factory->CreateToastNotification(xmlDoc, &toast);
    if(FAILED(hr)) goto cleanup;

    hr = WindowsCreateString(TOAST_AUMID, (UINT32)wcslen(TOAST_AUMID), &hAumid);
    if(FAILED(hr)) goto cleanup;

    hr = manager->CreateToastNotifierWithId(hAumid, &notifier);
    if(FAILED(hr)) goto cleanup;

    hr = notifier->Show(toast);
    if(FAILED(hr)) goto cleanup;

    ok = true;

cleanup:
    g_toast_last_hr = ok ? S_OK : hr;
    if(notifier) notifier->Release();
    if(toast) toast->Release();
    if(factory) factory->Release();
    if(xmlIO) xmlIO->Release();
    if(xmlDoc) xmlDoc->Release();
    if(manager) manager->Release();
    if(hAumid) WindowsDeleteString(hAumid);
    if(hToastClass) WindowsDeleteString(hToastClass);
    if(hXml) WindowsDeleteString(hXml);
    if(hManager) WindowsDeleteString(hManager);
    if(xml) free(xml);
    if(do_uninit) RoUninitialize();

    return ok;
}
