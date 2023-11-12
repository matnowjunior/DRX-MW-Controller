#include <windows.h>
#include <stdlib.h>
#include <string.h>
#include <tchar.h>
#include <commdlg.h>
#include <string>

#include "resource.h"

#include "ValuesSelector.h"
#include "globals.h"

#include <string>
#include <codecvt>
#include <locale>
#include <vector>

// Global variables

// The main window class name.
static TCHAR szWindowClass[] = _T("DesktopApp");

// The string that appears in the application's title bar.
static TCHAR szTitle[] = _T("DRX MW Controller");

// Stored instance handle for use in Win32 API calls such as FindResource
HINSTANCE hInst;

HWND hEditFilePath;
HWND hBrowseButton;
HWND hProceedButton;
HWND hFirstCoordsOfLegend;
HWND hEndCoordsOfLegend;
HWND hTestButton;

HWND hLegendButton;
HWND hSecondaryWindow;

HWND hCheckboxBak;
HWND hCheckboxOverride;

#define ID_BUTTON 1001

std::string stringPath;
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
        600, 250,
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
HPEN hWhitePen = CreatePen(PS_SOLID, 1, RGB(255, 255, 255));  // Global brush for button border
std::string WStringConverter(const std::wstring& s)
{
	std::wstring_convert<std::codecvt_utf8<wchar_t>> conv;
	return conv.to_bytes(s);
}
std::wstring GetEditControlText(HWND hWndEdit) {
    // Get the length of the text
    int length = GetWindowTextLength(hWndEdit);
    if (length == 0) {
        return L"";  // No text in the edit control
    }

    // Create a vector with the right size (+1 for null-terminator)
    std::vector<wchar_t> buffer(length + 1);

    // Retrieve the text
    GetWindowText(hWndEdit, &buffer[0], length + 1);

    // Convert the buffer to a wstring
    std::wstring text(&buffer[0]);

    return text;
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
    PAINTSTRUCT ps;
    HDC hdc;
    TCHAR greeting[] = _T("SWDBEFL, Test interfacu graficznego!");

    switch (message)
    {
    
    case WM_CREATE:
    {
        hEditFilePath = CreateWindowEx(WS_EX_CLIENTEDGE, L"EDIT", L"",
            WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL,
            10, 30, 250, 25, hWnd, NULL, GetModuleHandle(NULL), NULL);

        hBrowseButton = CreateWindow(L"BUTTON", L"Przegladaj",
            WS_CHILD | WS_VISIBLE | BS_OWNERDRAW,
            270, 30, 95, 25, hWnd, (HMENU)1, GetModuleHandle(NULL), NULL);

        hLegendButton = CreateWindow(L"BUTTON", L"Wprowadz dane legendy",
            WS_CHILD | WS_VISIBLE | BS_OWNERDRAW,
            400, 30, 150, 25, hWnd, (HMENU)2, GetModuleHandle(NULL), NULL);

        hProceedButton = CreateWindow(L"BUTTON", L"Zatwierdz",
            WS_CHILD | WS_VISIBLE | BS_OWNERDRAW,
            235, 150, 100, 25, hWnd, (HMENU)3, GetModuleHandle(NULL), NULL);

        // Create the checkboxes
        hCheckboxBak = CreateWindowEx(
            0, L"BUTTON", L"Stworz kopie zapasowa (.bak)",
            WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_CHECKBOX,
            10, 75, 220, 35, hWnd, (HMENU)101,
            (HINSTANCE)GetWindowLongPtr(hWnd, GWLP_HINSTANCE), NULL);
        hCheckboxOverride = CreateWindowEx(
            0, L"BUTTON", L"Nadpisz stary plik",
            WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_CHECKBOX,
            10, 100, 185, 35, hWnd, (HMENU)102,
            (HINSTANCE)GetWindowLongPtr(hWnd, GWLP_HINSTANCE), NULL);

        // Subclass the checkbox
        //SetWindowSubclass(hCheckbox, CheckboxSubclassProc, 0, 0);
        break;
    }
    case WM_COMMAND:
        switch ((LOWORD(wParam))) {
        case 1: {
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
            break; }
        case 2: {
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
                WS_OVERLAPPEDWINDOW, CW_USEDEFAULT, CW_USEDEFAULT, 180, 250, NULL, NULL, GetModuleHandle(NULL), NULL);

            ShowWindow(hSecondaryWindow, SW_SHOW);
            break; }
		case 3: {
            bool legendDataCorrect = true;
            for (int i = 0; i < 4; i++)
            {
                if (labelValues[i].empty())
                {
					legendDataCorrect = false;
                    MessageBox(
                        hWnd,                   // Parent window, if available, or NULL
                        _T("Nie wpisano wartosci legendy lub danych!"), // Message to be displayed
                        _T("Blad"),             // Title of the message box
                        MB_ICONERROR | MB_OK    // Style of the message box
                    );
					break;
				}
            }
            if (legendDataCorrect && sheetName.empty())
            {
                legendDataCorrect = false;
                MessageBox(
                    hWnd,                   // Parent window, if available, or NULL
                    _T("Nie wybrano arkusza!"), // Message to be displayed
                    _T("Blad"),             // Title of the message box
                    MB_ICONERROR | MB_OK    // Style of the message box
                );
                break;
            }
            InvalidateRect((HWND)lParam, NULL, TRUE);  // force button to redraw
            Convert(WStringConverter(GetEditControlText(hEditFilePath)), WStringConverter(sheetName), WStringConverter(labelValues[0]), WStringConverter(labelValues[1]), WStringConverter(labelValues[2]), WStringConverter(labelValues[3]));
            break; }
        case 101:
            if (HIWORD(wParam) == BN_CLICKED) {
				BOOL checked = IsDlgButtonChecked(hWnd, 101);
                if (checked) {
					// The checkbox is checked
					CheckDlgButton(hWnd, 101, BST_UNCHECKED);
				}
                else {
					// The checkbox is unchecked
					CheckDlgButton(hWnd, 101, BST_CHECKED);
				}
			}
			break;
        case 102:
            if (HIWORD(wParam) == BN_CLICKED) {
                BOOL checked = IsDlgButtonChecked(hWnd, 102);
                if (checked) {
                    // The checkbox is checked
                    CheckDlgButton(hWnd, 102, BST_UNCHECKED);
                }
                else {
                    // The checkbox is unchecked
                    CheckDlgButton(hWnd, 102, BST_CHECKED);
                }
            }
            break;
        }
    case WM_CTLCOLORBTN:
        if ((HWND)lParam == hBrowseButton || (HWND)lParam == hLegendButton || (HWND)lParam == hProceedButton)
        {
            HDC hdcButton = (HDC)wParam;
            SetTextColor(hdcButton, RGB(255, 255, 255)); // Text Color: White

            if (hButtonBkGnd == NULL)
                hButtonBkGnd = CreateSolidBrush(RGB(50, 50, 50)); // Brush for Gray Background

            return (INT_PTR)hButtonBkGnd;
        }
        break;
    case WM_CTLCOLORSTATIC:
    {
        HDC hdcStatic = (HDC)wParam;
        SetTextColor(hdcStatic, RGB(255, 255, 255)); // Set text color to white
        SetBkMode(hdcStatic, TRANSPARENT); // Set background mode to transparent

        // Return a hollow brush to maintain a transparent background
        static HBRUSH hBrushTransparent = (HBRUSH)GetStockObject(HOLLOW_BRUSH);
        return (LRESULT)hBrushTransparent;
    }
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
        if (lpDIS->CtlID == 3) // Button's ID
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
            DrawText(lpDIS->hDC, L"Zatwierdz", -1, &lpDIS->rcItem, DT_CENTER | DT_SINGLELINE | DT_VCENTER);
        }
        break;
    }
    
    case WM_ERASEBKGND: // For background not to dissapear after minimalizing
    {
        HDC hdc = (HDC)wParam;
        RECT rect;
        GetClientRect(hWnd, &rect); // Get the dimensions of the window
        HBRUSH hbrBackground = CreateSolidBrush(RGB(0x24, 0x24, 0x24)); // Or use a global brush if you have one
        FillRect(hdc, &rect, hbrBackground);
        DeleteObject(hbrBackground); // If you created a new brush
        return 1; // Return non-zero to indicate that you handled the message
    }
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
