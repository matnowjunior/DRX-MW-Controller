#include <string>
#include <iostream>
#include <xlnt/xlnt.hpp>
#include <fstream>
#include <filesystem>
#include <cstring>
#include <cwchar>
#include <windows.h>
#include "globals.h"
#include <string>
#include <codecvt>
#include <locale>
#include "libxl.h"

using namespace std;
using namespace libxl;

int coordinates[4];

vector<Color> colors;

int titleToNumber(string s)
{
    int r = 0;
    for (int i = 0; i < s.length(); i++)
    {
        r = r * 26 + s[i] - 64;
    }
    return r;
}

pair <int, int> signs_numbers_separately(wstring cell_input)
{
    string letters, digits;
    for (char c : cell_input) {
        if (isalpha(c)) {
            letters += c;  //adding letter to letter variable
        }
        else if (isdigit(c)) {
            digits += c;  //adding number to digits variable
        }
    }

    return make_pair(titleToNumber(letters), stoi(digits));

}

void get_cell(wstring first_cell, wstring second_cell)
{

    coordinates[0] = signs_numbers_separately(first_cell).first;
    coordinates[1] = signs_numbers_separately(first_cell).second;

    coordinates[2] = signs_numbers_separately(second_cell).first;
    coordinates[3] = signs_numbers_separately(second_cell).second;

    if (coordinates[2] < coordinates[0])
        swap(coordinates[2], coordinates[0]);
    if (coordinates[3] < coordinates[1])
        swap(coordinates[3], coordinates[1]);
}

void setCellFillColor(Sheet* sheet, int row, int col, const Color& fillColor) {
    if (sheet) {
        Format* cellFormat = sheet->cellFormat(row, col);
        cellFormat->setPatternBackgroundColor(fillColor);
        sheet->writeBlank(row, col, cellFormat);
    }
}

void getCellColor(Sheet* sheet, int row, int col) {
    //Format* format = sheet->cellFormat(row, col);
    //int colorValue = format->fillPattern();

    //// Store the color in the vector
    //colors.push_back(format->patternForegroundColor());
}

LPCWSTR stringToLPCWSTR(const std::string& str) {
    // Convert std::string to std::wstring
    std::wstring_convert<std::codecvt_utf8_utf16<wchar_t>> converter;
    std::wstring wideStr = converter.from_bytes(str);

    // Return LPCWSTR (pointer to the wide string)
    return wideStr.c_str();
}
LPCWSTR charToLPCWSTR(const char* charArray) {
    // Calculate the length of the result wide string
    int len = MultiByteToWideChar(CP_UTF8, 0, charArray, -1, NULL, 0);

    // Allocate buffer for the wide string
    wchar_t* wString = new wchar_t[len];

    // Perform the conversion
    MultiByteToWideChar(CP_UTF8, 0, charArray, -1, wString, len);

    return wString;
}

void getCellColor(Book* book, Sheet* sheet, int i, int j)
{
    Format* format1 = book->addFormat();

    format1 = sheet->cellFormat(i, j);
    format1->setFillPattern(FILLPATTERN_SOLID);
    colors.push_back(format1->patternForegroundColor());
}

void setCellFillColor(Book* book, Sheet* sheet, int i, int j, Color color)
{
    Format* format2 = book->addFormat();
    Font* font = book->addFont();


    double number = sheet->readNum(i, j);

    font = sheet->cellFormat(i, j)->font();
    format2->setFont(font);
    format2->setPatternForegroundColor(color);
    format2->setFillPattern(FILLPATTERN_SOLID);
    sheet->writeBlank(i, j, format2);
    sheet->writeNum(i, j, number, format2);
}

void Convert(wstring path, int sheet1, wstring cl1, wstring cl2, wstring cd1, wstring cd2)
{
    Book* book = xlCreateXMLBook();
    book->setRgbMode(true);

    book->load(L"D:/excel/DRX MW CONTROLLER/DRX-MW-Controller/heh.xlsx");

    int activeSheetIndex = book->activeSheet();
    Sheet* sheet = book->getSheet(activeSheetIndex);
    // FillPattern fillPatterns = FILLPATTERN_SOLID;

     //Format* format2 = book->addFormat();
     //format2->setFillPattern(fillPatterns);
     //format2->setPatternForegroundColor(book->colorPack(50, 50, 50));*/

     /*getCellColor(book, sheet, 1, 4);
     getCellColor(book, sheet, 1, 5);

     setCellFillColor(book, sheet, 1, 0, colors[0]);
     setCellFillColor(book, sheet, 1, 1, colors[1]);*/

    get_cell(cl1, cl2);
    double* min_cell_values = new double[(coordinates[2] - coordinates[0]) + 2];
    double* max_cell_values = new double[(coordinates[2] - coordinates[0]) + 2];

    int col_counter;
    int index = 0;
    int length = coordinates[2] - coordinates[0];


    //pętla legendy
    for (int i = coordinates[0]; i < (coordinates[2] + 1); i++)
    {
        col_counter = 1;
        for (int j = coordinates[1]; j < (coordinates[3] + 1); j++)
        {
            switch (col_counter)
            {
            case 1:
            {
                getCellColor(book, sheet, j - 1, i - 1);
                break;
            }
            case 2:
            {
                min_cell_values[index] = sheet->readNum(j - 1, i - 1);
                //cout << "min: " << min_cell_values[i] << endl;
                break;
            }
            case 3:
            {
                max_cell_values[index] = sheet->readNum(j - 1, i - 1);
                //cout << "max:" << max_cell_values[i] << endl;
                break;
            }
            default:
                break;
            }
            col_counter++;

        }
        index++;

    }

    //pętla wartości
    get_cell(cd1, cd2);

    for (int i = coordinates[0]; i < coordinates[2] + 1; i++)
    {
        for (int j = coordinates[1]; j < coordinates[3] + 1; j++)
        {

            //cout << to_string(cell_get(j, i, ws).value<int>()) << " ";

            for (int w = 0; w < length + 1; w++)
            {
                //cout << "w: " << w << endl;
                //cout << to_string(min_cell_values[w]) << endl;
                //cout << to_string(max_cell_values[w]) << endl;
                int cellValue = sheet->readNum(i, j);

                // Check if the cell has a value (not equal to -1 or any other indicator)
                if (cellValue != -1)
                {
                    if (sheet->readNum(j - 1, i - 1) > min_cell_values[w] && sheet->readNum(j - 1, i - 1) <= max_cell_values[w])
                    {
                        int numer = sheet->readNum(j - 1, i - 1);
                        setCellFillColor(book, sheet, j - 1, i - 1, colors[w]);

                        break;
                    }
                }
                                
            }

        }

    }

    book->save(L"example.xlsx");
    book->release();
}
