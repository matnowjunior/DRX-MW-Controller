#include <string.h>
#include <windows.h>
#include <string.h>
#include <string>      // Required for std::wstring
#include <regex>       // Required for std::wregex and std::regex_match
#include <cwctype>  // Required for towupper

#include "resource.h"
#include "ValuesSelector.h"

HWND hApply;
HWND hCellData[4];
HBRUSH hVSButtonBkGnd;
HPEN hWhitePen2 = NULL;

HFONT hArialFont = CreateFont(16,          // Height of font
    0,           // Average character width
    0,           // Angle of escapement
    0,           // Base-line orientation angle
    FW_BOLD,   // Font weight
    FALSE,       // Italic
    FALSE,       // Underline
    0,           // Strikeout
    ANSI_CHARSET,// Character set identifier
    OUT_TT_PRECIS,// Output precision
    CLIP_DEFAULT_PRECIS, // Clipping precision
    ANTIALIASED_QUALITY, // Output quality
    FF_DONTCARE | DEFAULT_PITCH, // Family and pitch
    L"Arial");   // Font name

LRESULT CALLBACK EditSubclassProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam) {
    switch (uMsg) {
    case WM_CHAR: {
        wchar_t charToProcess = static_cast<wchar_t>(wParam);

        if (iswlower(charToProcess)) {
            charToProcess = towupper(charToProcess);
        }

        if (iswalpha(charToProcess) || iswdigit(charToProcess) || charToProcess == VK_BACK) { // Check if it's a letter, digit, or backspace
            WCHAR buffer[1024];
            GetWindowText(hWnd, buffer, 1024);
            std::wstring currentText = buffer;
            std::wstring newText = currentText + charToProcess;
            std::wregex pattern(L"^[A-Za-z]+[0-9]*$");
            if (charToProcess != VK_BACK && !std::regex_match(newText, pattern)) {
                MessageBeep(MB_ICONERROR); // Play error sound
                return 0; // Do not process this character
            }
        }
        else {
            MessageBeep(MB_ICONERROR); // Play error sound
            return 0; // Reject special characters
        }
        wParam = static_cast<WPARAM>(charToProcess); // Update the wParam with possibly changed character
    } break;
    }
    // Call the original EDIT control window procedure to process other messages and default behavior
    return CallWindowProc((WNDPROC)GetWindowLongPtr(hWnd, GWLP_USERDATA), hWnd, uMsg, wParam, lParam);
}


LRESULT CALLBACK SecondaryWindowProcedure(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam) {
    switch (uMsg) {
    case WM_PAINT: {
        PAINTSTRUCT ps;
        HDC hdc = BeginPaint(hWnd, &ps);

        HFONT hOldFont = (HFONT)SelectObject(hdc, hArialFont);
        SetTextColor(hdc, RGB(255, 255, 255)); // White text
        SetBkMode(hdc, TRANSPARENT);           // Transparent background
        // Begin application-specific layout section.
        TextOut(hdc, 25, 5, L"Pozycja Legendy:", wcslen(L"Pozycja Legendy:"));
        TextOut(hdc, 30, 60, L"Pozycja Danych:", wcslen(L"Pozycja Danych:"));
        // End application-specific layout section.

        EndPaint(hWnd, &ps);
        // Restore original font
        SelectObject(hdc, hOldFont);

        // When you're done using the font, remember to delete it to free resources:
        DeleteObject(hArialFont);
    } break;
    case WM_CREATE: {
        // Create the button
        if (!hWhitePen2)
            hWhitePen2 = CreatePen(PS_SOLID, 1, RGB(255, 255, 255));
        if (!hVSButtonBkGnd)
            hVSButtonBkGnd = CreateSolidBrush(RGB(50, 50, 50));
        hApply = CreateWindow(L"BUTTON", L"",
            WS_CHILD | WS_VISIBLE | BS_OWNERDRAW,
            35, 120, 100, 25, hWnd, (HMENU)401, GetModuleHandle(NULL), NULL);
        hCellData[0] = CreateWindow(L"EDIT", L"",
            WS_CHILD | WS_VISIBLE | WS_BORDER | ES_LEFT,
            15, 30, 50, 25, hWnd, (HMENU)410, GetModuleHandle(NULL), NULL);
        hCellData[1] = CreateWindow(L"EDIT", L"",
            WS_CHILD | WS_VISIBLE | WS_BORDER | ES_LEFT,
            95, 30, 50, 25, hWnd, (HMENU)411, GetModuleHandle(NULL), NULL);
        hCellData[2] = CreateWindow(L"EDIT", L"",
            WS_CHILD | WS_VISIBLE | WS_BORDER | ES_LEFT,
            15, 80, 50, 25, hWnd, (HMENU)412, GetModuleHandle(NULL), NULL);
        hCellData[3] = CreateWindow(L"EDIT", L"",
            WS_CHILD | WS_VISIBLE | WS_BORDER | ES_LEFT,
            95, 80, 50, 25, hWnd, (HMENU)413, GetModuleHandle(NULL), NULL);
        for (int i = 0; i < 4; i++)
        {
            SetWindowLongPtr(hCellData[i], GWLP_USERDATA, (LONG_PTR)SetWindowLongPtr(hCellData[i], GWLP_WNDPROC, (LONG_PTR)EditSubclassProc));
        }
        
        InvalidateRect(hApply, NULL, TRUE);  // force button to redraw
    }break;
        
    case WM_CTLCOLORBTN:
        if ((HWND)lParam == hApply)
        {
            HDC hdcButton = (HDC)wParam;
            SetTextColor(hdcButton, RGB(255, 255, 255)); // Text Color: White

            return (INT_PTR)hVSButtonBkGnd;
        }
        break;

    case WM_DRAWITEM:
    {
        LPDRAWITEMSTRUCT lpDIS = (LPDRAWITEMSTRUCT)lParam;
        if (lpDIS->CtlID == 401) // Button's ID
        {

            // Fill the button with gray color
            FillRect(lpDIS->hDC, &lpDIS->rcItem, hVSButtonBkGnd);

            // Draw the white border
            HPEN hOldPen = (HPEN)SelectObject(lpDIS->hDC, hWhitePen2);
            Rectangle(lpDIS->hDC, lpDIS->rcItem.left, lpDIS->rcItem.top, lpDIS->rcItem.right, lpDIS->rcItem.bottom);
            SelectObject(lpDIS->hDC, hOldPen);

            // Draw the button text
            SetTextColor(lpDIS->hDC, RGB(255, 255, 255));
            SetBkMode(lpDIS->hDC, TRANSPARENT);
            DrawText(lpDIS->hDC, L"Zatwierdz", -1, &lpDIS->rcItem, DT_CENTER | DT_SINGLELINE | DT_VCENTER);
            return TRUE;  // message processed
        }
        break;
    }
    case WM_COMMAND:
        if (LOWORD(wParam) == 401) {  // Button's ID
            InvalidateRect((HWND)lParam, NULL, TRUE);  // force button to redraw
        }
        break;
    case WM_CLOSE:
        DestroyWindow(hWnd);
        return 0;

    default:
        return DefWindowProc(hWnd, uMsg, wParam, lParam);
    }

    return 0;
}
