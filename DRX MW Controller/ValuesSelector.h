// SecondaryWindow.h

#ifndef VALUES_SELECTOR_H
#define VALUES_SELECTOR_H

#include <windows.h>

// Global variables declaration
extern HBRUSH hEditBkGnd;
extern HBRUSH hButtonBkGnd;
extern HPEN hWhitePen;

// Function prototype
LRESULT CALLBACK SecondaryWindowProcedure(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam);

#endif // VALUES_SELECTOR_H
