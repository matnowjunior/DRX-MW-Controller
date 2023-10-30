#include <windows.h>
#include <stdlib.h>
#include <string.h>
#include <tchar.h>
#include <commdlg.h>
#include <string>

#include "resource.h"

#include "ValuesSelector.h"


// Global variables

// The main window class name.
static TCHAR szWindowClass[] = _T("DesktopApp");

// The string that appears in the application's title bar.
static TCHAR szTitle[] = _T("DRX MW Controller");

// Stored instance handle for use in Win32 API calls such as FindResource
HINSTANCE hInst;

HWND hEditFilePath;
HWND hBrowseButton;
HWND hFirstCoordsOfLegend;
HWND hEndCoordsOfLegend;

HWND hLegendButton;
HWND hSecondaryWindow;


#define ID_BUTTON 1001

// Forward declarations of functions included in this code module:
LRESULT CALLBACK WndProc(HWND, UINT, WPARAM, LPARAM);

int WINAPI WinMain(
    _In_ HINSTANCE hInstance,
    _In_opt_ HINSTANCE hPrevInstance,
    _In_ LPSTR     lpCmdLine,
    _In_ int       nCmdShow
)
{
    WNDCLASSEX wcex;

    wcex.cbSize = sizeof(WNDCLASSEX);
    wcex.style = CS_HREDRAW | CS_VREDRAW;
    wcex.lpfnWndProc = WndProc;
    wcex.cbClsExtra = 0;
    wcex.cbWndExtra = 0;
    wcex.hInstance = hInstance;
    wcex.hIcon = LoadIcon(wcex.hInstance, IDI_APPLICATION);
    wcex.hCursor = LoadCursor(NULL, IDC_ARROW);
    HBRUSH hCustomBrush = CreateSolidBrush(RGB(0x24, 0x24, 0x24));
    wcex.hbrBackground = hCustomBrush;
    wcex.lpszMenuName = NULL;
    wcex.lpszClassName = szWindowClass;
    wcex.hIconSm = LoadIcon(wcex.hInstance, MAKEINTRESOURCE(IDI_ICON1));

    if (!RegisterClassEx(&wcex))
    {
        MessageBox(NULL,
            _T("Call to RegisterClassEx failed!"),
            _T("Windows Desktop Guided Tour"),
            NULL);

        return 1;
    }

    // Store instance handle in our global variable
    hInst = hInstance;

    // The parameters to CreateWindowEx explained:
    // WS_EX_OVERLAPPEDWINDOW : An optional extended window style.
    // szWindowClass: the name of the application
    // szTitle: the text that appears in the title bar
    // WS_OVERLAPPEDWINDOW: the type of window to create
    // CW_USEDEFAULT, CW_USEDEFAULT: initial position (x, y)
    // 500, 100: initial size (width, length)
    // NULL: the parent of this window
    // NULL: this application does not have a menu bar
    // hInstance: the first parameter from WinMain
    // NULL: not used in this application
    HWND hWnd = CreateWindowEx(
        WS_EX_OVERLAPPEDWINDOW,
        szWindowClass,
        szTitle,
        WS_OVERLAPPEDWINDOW,
        CW_USEDEFAULT, CW_USEDEFAULT,
        1000, 500,
        NULL,
        NULL,
        hInstance,
        NULL
    );

    if (!hWnd)
    {
        MessageBox(NULL,
            _T("Call to CreateWindow failed!"),
            _T("Windows Desktop Guided Tour"),
            NULL);

        return 1;
    }

    // The parameters to ShowWindow explained:
    // hWnd: the value returned from CreateWindow
    // nCmdShow: the fourth parameter from WinMain
    ShowWindow(hWnd,
        nCmdShow);
    UpdateWindow(hWnd);
    DeleteObject(hCustomBrush);
    // Main message loop:
    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0))
    {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    return (int)msg.wParam;
}

//  FUNCTION: WndProc(HWND, UINT, WPARAM, LPARAM)
//
//  PURPOSE:  Processes messages for the main window.
//
//  WM_PAINT    - Paint the main window
//  WM_DESTROY  - post a quit message and return


HBRUSH hEditBkGnd = NULL; // Global brush for edit control background
HBRUSH hButtonBkGnd = NULL;
HPEN hWhitePen = NULL; // Global brush for button border

LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
    PAINTSTRUCT ps;
    HDC hdc;
    TCHAR greeting[] = _T("SWDBEFL, Test interfacu graficznego!");

    switch (message)
    {
    
    case WM_CREATE:
        hEditFilePath = CreateWindowEx(WS_EX_CLIENTEDGE, L"EDIT", L"",
            WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL,
            10, 30, 250, 25, hWnd, NULL, GetModuleHandle(NULL), NULL);

        hBrowseButton = CreateWindow(L"BUTTON", L"Przegladaj",
            WS_CHILD | WS_VISIBLE | BS_OWNERDRAW, // Add BS_OWNERDRAW style
            270, 30, 95, 25, hWnd, (HMENU)1, GetModuleHandle(NULL), NULL);

        hLegendButton = CreateWindow(L"BUTTON", L"Wprowadz dane legendy",
            WS_CHILD | WS_VISIBLE | BS_OWNERDRAW,
            700, 10, 200, 25, hWnd, (HMENU)2, GetModuleHandle(NULL), NULL);

        SendMessage(hFirstCoordsOfLegend, EM_SETLIMITTEXT, (WPARAM)5, 0);
        SendMessage(hEndCoordsOfLegend, EM_SETLIMITTEXT, (WPARAM)5, 0);
        SetWindowText(hFirstCoordsOfLegend, L"D10");
        SetWindowText(hEndCoordsOfLegend, L"N12");
        break;

    case WM_COMMAND:
        if (LOWORD(wParam) == 1) {  // Button's ID
            OPENFILENAME ofn = { 0 };
            WCHAR filePath[MAX_PATH] = { 0 }; // Using WCHAR for wide characters

            ofn.lStructSize = sizeof(ofn);
            ofn.hwndOwner = hWnd;
            ofn.lpstrFile = filePath;
            ofn.nMaxFile = sizeof(filePath) / sizeof(WCHAR); // Adjusted for wide characters
            ofn.lpstrFilter = L"Pliki Excel (*.xlsx)\0*.xlsx\0Wszystko\0*.*\0";
            ofn.nFilterIndex = 1;
            ofn.lpstrFileTitle = NULL;
            ofn.nMaxFileTitle = 0;
            ofn.lpstrInitialDir = NULL;
            ofn.Flags = OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST;

            if (GetOpenFileName(&ofn)) {
                SetWindowText(hEditFilePath, filePath); // filePath is now wide character string
            }
            InvalidateRect((HWND)lParam, NULL, TRUE);  // force button to redraw
        }
        if (LOWORD(wParam) == 2) {
            WNDCLASSEX wcCheck;
            if (!GetClassInfoEx(GetModuleHandle(NULL), L"SecondaryWindowClass", &wcCheck)) {
                // If not, register the class
                WNDCLASSEX wc = { 0 };
                wc.cbSize = sizeof(WNDCLASSEX);
                wc.lpfnWndProc = SecondaryWindowProcedure;
                wc.hInstance = GetModuleHandle(NULL);
                wc.hCursor = LoadCursor(NULL, IDC_ARROW);
                wc.hIcon = LoadIcon(wc.hInstance, IDI_APPLICATION);
                wc.hIconSm = LoadIcon(wc.hInstance, MAKEINTRESOURCE(IDI_ICON1));
                HBRUSH hCustomBrush = CreateSolidBrush(RGB(0x24, 0x24, 0x24));
                wc.hbrBackground = hCustomBrush;
                wc.lpszClassName = L"SecondaryWindowClass";

                if (!RegisterClassEx(&wc)) {
                    MessageBox(NULL, L"Failed to register secondary window class.", L"Error", MB_ICONERROR);
                    break;
                }
            }

            hSecondaryWindow = CreateWindowEx(0, L"SecondaryWindowClass", L"Text Inputs",
                WS_OVERLAPPEDWINDOW, CW_USEDEFAULT, CW_USEDEFAULT, 200, 200, NULL, NULL, GetModuleHandle(NULL), NULL);

            ShowWindow(hSecondaryWindow, SW_SHOW);
        }
        break;
    case WM_CTLCOLORBTN:
        if ((HWND)lParam == hBrowseButton || (HWND)lParam == hLegendButton)
        {
            HDC hdcButton = (HDC)wParam;
            SetTextColor(hdcButton, RGB(255, 255, 255)); // Text Color: White

            if (hButtonBkGnd == NULL)
                hButtonBkGnd = CreateSolidBrush(RGB(50, 50, 50)); // Brush for Gray Background

            return (INT_PTR)hButtonBkGnd;
        }
        break;

    case WM_DRAWITEM:
    {
        LPDRAWITEMSTRUCT lpDIS = (LPDRAWITEMSTRUCT)lParam;
        if (lpDIS->CtlID == 1) // Button's ID
        {
            if (hWhitePen == NULL)
                hWhitePen = CreatePen(PS_SOLID, 1, RGB(255, 255, 255)); // White pen for border

            // Fill the button with gray color
            FillRect(lpDIS->hDC, &lpDIS->rcItem, hButtonBkGnd);

            // Draw the white border
            HPEN hOldPen = (HPEN)SelectObject(lpDIS->hDC, hWhitePen);
            Rectangle(lpDIS->hDC, lpDIS->rcItem.left, lpDIS->rcItem.top, lpDIS->rcItem.right, lpDIS->rcItem.bottom);
            SelectObject(lpDIS->hDC, hOldPen);

            // Draw the button text
            SetTextColor(lpDIS->hDC, RGB(255, 255, 255)); // White text
            SetBkMode(lpDIS->hDC, TRANSPARENT);
            DrawText(lpDIS->hDC, L"Przegladaj", -1, &lpDIS->rcItem, DT_CENTER | DT_SINGLELINE | DT_VCENTER);
        }
        if (lpDIS->CtlID == 2) // Button's ID
        {
            if (hWhitePen == NULL)
                hWhitePen = CreatePen(PS_SOLID, 1, RGB(255, 255, 255)); // White pen for border

            // Fill the button with gray color
            FillRect(lpDIS->hDC, &lpDIS->rcItem, hButtonBkGnd);

            // Draw the white border
            HPEN hOldPen = (HPEN)SelectObject(lpDIS->hDC, hWhitePen);
            Rectangle(lpDIS->hDC, lpDIS->rcItem.left, lpDIS->rcItem.top, lpDIS->rcItem.right, lpDIS->rcItem.bottom);
            SelectObject(lpDIS->hDC, hOldPen);

            // Draw the button text
            SetTextColor(lpDIS->hDC, RGB(255, 255, 255)); // White text
            SetBkMode(lpDIS->hDC, TRANSPARENT);
            DrawText(lpDIS->hDC, L"Podaj dane pliku", -1, &lpDIS->rcItem, DT_CENTER | DT_SINGLELINE | DT_VCENTER);
        }
    }
    break;

    case WM_DESTROY:
        if (hButtonBkGnd)
            DeleteObject(hButtonBkGnd); // Cleanup the brush
        if (hWhitePen)
            DeleteObject(hWhitePen); // Cleanup the pen
        UnregisterClass(L"SecondaryWindowClass", GetModuleHandle(NULL));
        DeleteObject(hEditBkGnd);   // delete the edit control background brush
        DeleteObject(hButtonBkGnd); // delete the button background brush
        PostQuitMessage(0);
        break;

    default:
        return DefWindowProc(hWnd, message, wParam, lParam);
        break;
    }

    return 0;
}

//using namespace std;
//int main()
//{
//    HINSTANCE hInstance = GetModuleHandle(NULL);
//    return WinMain(hInstance, NULL, GetCommandLineA(), SW_SHOW);
//    xlnt::workbook wb;
//    wb.load("D:/Test.xlsx");
//    auto ws = wb.active_sheet();
//    std::clog << "Processing spread sheet" << std::endl;
//    for (auto row : ws.rows(false))
//    {
//        for (auto cell : row)
//        {
//            xlnt::cell_reference cellRef(cell.column_index(), cell.row());
//            std::clog << cellRef.to_string() << ": " << cell.to_string() << std::endl;
//        }
//    }
//    std::clog << "Processing complete" << std::endl;
//    return 0;
//}