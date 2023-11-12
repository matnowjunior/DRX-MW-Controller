// globals.h

#ifndef GLOBALS_H
#define GLOBALS_H

#include <windows.h>
#include <string>


// Declare the array as extern so other files know of its existence
extern std::wstring labelValues[4];
extern std::wstring sheetName;

void Convert(std::string path, std::string sheet, std::string cl1, std::string cl2, std::string cd1, std::string cd2);

#endif // GLOBALS_H
