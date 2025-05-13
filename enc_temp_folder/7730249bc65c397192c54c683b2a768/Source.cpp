#pragma comment(lib, "Shlwapi")

// nuget で Microsoft.Web.WebView2 と Microsoft.Windows.ImplementationLibrary をインストールする

#include <windows.h>
#include <wil/com.h>
#include <WebView2.h>
#include <string>
#include <iostream>
#include <wrl/event.h>
#include <shlobj_core.h>
#include "Shlwapi.h"

using namespace Microsoft::WRL;
TCHAR szClassName[] = TEXT("Window");

wil::com_ptr<ICoreWebView2Controller> g_controller;
wil::com_ptr<ICoreWebView2> g_webView;
HWND hOutputEdit;

class ScriptCallback : public RuntimeClass<RuntimeClassFlags<ClassicCom>, ICoreWebView2ExecuteScriptCompletedHandler> {
public:
    HRESULT STDMETHODCALLTYPE Invoke(HRESULT errorCode, LPCWSTR resultObjectAsJson) override {
        if (SUCCEEDED(errorCode)) {
            std::wstring result(resultObjectAsJson);
            if (result.length() >= 2 && result.front() == L'"' && result.back() == L'"') {
                result = result.substr(1, result.length() - 2);
            }
            SetWindowText(hOutputEdit, result.c_str());
        }
        else {
            SetWindowText(hOutputEdit, L"JavaScript 実行エラー");
        }
        return S_OK;
    }
};

void InitializeWebView(HWND hwnd, LPWSTR lpszJavaScript) {
    CreateCoreWebView2EnvironmentWithOptions(nullptr, 0, nullptr,
        Callback<ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler>(
            [hwnd, lpszJavaScript](HRESULT result, ICoreWebView2Environment* env) -> HRESULT {
                env->CreateCoreWebView2Controller(hwnd,
                    Callback<ICoreWebView2CreateCoreWebView2ControllerCompletedHandler>(
                        [hwnd, lpszJavaScript](HRESULT result, ICoreWebView2Controller* controller) -> HRESULT {
                            if (FAILED(result)) {
                                return E_FAIL;
                            }
                            g_controller = controller;
                            g_controller->get_CoreWebView2(&g_webView);
                            g_webView->Navigate(L"about:blank");
                            g_webView->add_NavigationCompleted(
                                Callback<ICoreWebView2NavigationCompletedEventHandler>(
                                    [lpszJavaScript](ICoreWebView2* sender, ICoreWebView2NavigationCompletedEventArgs* args) -> HRESULT {
                                        BOOL isSuccess = FALSE;
                                        HRESULT hr = args->get_IsSuccess(&isSuccess);
                                        if (SUCCEEDED(hr) && isSuccess) {
                                            sender->ExecuteScript(lpszJavaScript, Make<ScriptCallback>().Get());
                                        }
                                        return S_OK;
                                    }).Get(), nullptr);
                            return S_OK;
                        }).Get());
                return S_OK;
            }).Get());
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    static HWND hButton;
    static HWND hEdit1;
    static LPWSTR lpszJavaScript;
    switch (msg) {
    case WM_CREATE:
        InitializeWebView(hWnd, lpszJavaScript);
        hEdit1 = CreateWindowEx(WS_EX_CLIENTEDGE, TEXT("EDIT"), L"JSON.stringify(1 + 2);", WS_VISIBLE | WS_CHILD | ES_MULTILINE | ES_AUTOHSCROLL | ES_AUTOVSCROLL, 0, 0, 0, 0, hWnd, 0, ((LPCREATESTRUCT)lParam)->hInstance, 0);
        hButton = CreateWindow(TEXT("BUTTON"), TEXT("Javascript実行"), WS_VISIBLE | WS_CHILD, 0, 0, 0, 0, hWnd, (HMENU)IDOK, ((LPCREATESTRUCT)lParam)->hInstance, 0);
        hOutputEdit = CreateWindowEx(WS_EX_CLIENTEDGE, TEXT("EDIT"), 0, WS_VISIBLE | WS_CHILD | ES_MULTILINE | ES_AUTOHSCROLL | ES_AUTOVSCROLL, 0, 0, 0, 0, hWnd, 0, ((LPCREATESTRUCT)lParam)->hInstance, 0);
        break;
    case WM_SIZE:
        MoveWindow(hEdit1, 0, 0, LOWORD(lParam), HIWORD(lParam) / 2 - 16, TRUE);
        MoveWindow(hButton, 0, HIWORD(lParam) / 2 - 16, 128, 32, TRUE);
        MoveWindow(hOutputEdit, 0, HIWORD(lParam) / 2 + 16, LOWORD(lParam), HIWORD(lParam) / 2 - 16, TRUE);
        break;
    case WM_COMMAND:
        if (LOWORD(wParam) == IDOK && g_webView)
        {
            if (lpszJavaScript) {
                GlobalFree(lpszJavaScript);
                lpszJavaScript = 0;
            }
            DWORD dwSize = GetWindowTextLength(hEdit1);
            lpszJavaScript = (LPWSTR)GlobalAlloc(0, sizeof(WCHAR) * (dwSize + 1));
            if (lpszJavaScript) {
                GetWindowText(hEdit1, lpszJavaScript, dwSize + 1);
                g_webView->ExecuteScript(lpszJavaScript, Make<ScriptCallback>().Get());
            }
        }
        break;
    case WM_DESTROY:
        g_webView.reset();
        g_controller.reset();
        if (lpszJavaScript) {
            GlobalFree(lpszJavaScript);
            lpszJavaScript = 0;
        }
        PostQuitMessage(0);
        break;
    default:
        return DefWindowProc(hWnd, msg, wParam, lParam);
    }
    return 0;
}

int WINAPI wWinMain(_In_ HINSTANCE hInstance,
    _In_opt_ HINSTANCE hPrevInstance,
    _In_ LPWSTR lpCmdLine,
    _In_ int nShowCmd)
{
    (void)CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    MSG msg;
    WNDCLASS wndclass = {
        CS_HREDRAW | CS_VREDRAW,
        WndProc,
        0,
        0,
        hInstance,
        0,
        LoadCursor(0,IDC_ARROW),
        (HBRUSH)(COLOR_WINDOW + 1),
        0,
        szClassName
    };
    RegisterClass(&wndclass);
    HWND hWnd = CreateWindow(
        szClassName,
        L"Javascript実行",
        WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN,
        CW_USEDEFAULT,
        0,
        CW_USEDEFAULT,
        0,
        0,
        0,
        hInstance,
        0
    );
    ShowWindow(hWnd, SW_SHOWDEFAULT);
    UpdateWindow(hWnd);
    while (GetMessage(&msg, 0, 0, 0))
    {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }
    CoUninitialize();
    return (int)msg.wParam;
}
