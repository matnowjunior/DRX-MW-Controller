#include <string>
#include <iostream>
#include <xlnt/xlnt.hpp>
#include <fstream>
#include <filesystem>

#include <windows.h>
#include "globals.h"
#include <string>
#include <codecvt>
#include <locale>
//define colors
#define RED "\033[31m"    
#define GREEN "\033[32m" 
#define RESET "\033[0m"

using namespace std;

//global variables
string announcement, color;
int coordinates[4];

xlnt::rgb_color get_cell_color_solid(const xlnt::cell& cell)
{
    xlnt::rgb_color color = cell.fill().pattern_fill().foreground().get().rgb();
    return color;
}

xlnt::cell cell_get(int x, int y, const xlnt::worksheet ws)
{

    xlnt::cell cell = ws.cell(x, y);
    return cell;
}

//function which fill specified cell with color
void color_fill_cell(const xlnt::cell& cell, const xlnt::rgb_color& rgb) {

    xlnt::cell cell1 = cell;
    cell1.fill(xlnt::fill::solid(rgb));
    cell.fill() = cell1.fill();

}

//string wybierzArkusz(xlnt::workbook& wb)
//{
//    xlnt::worksheet ws;
//    int num_sheets = wb.sheet_count();//returns number of workshhets in this workbook
//    cout << "Avaiable sheets:" << endl;
//    for (int i = 1; i <= num_sheets; i++) {
//        ws = wb.sheet_by_index(i - 1);//returns worksheet at given index
//        cout << i << ". " << ws.title() << endl;//returns title of the sheet
//    }
//
//    int wybor;
//    while (true) {
//        cout << "Choose sheet number: " << endl;
//        cin >> wybor;
//        if (wybor >= 1 && wybor <= num_sheets) {
//            xlnt::worksheet wybranyArkusz = wb.sheet_by_index(wybor - 1);
//            return wybranyArkusz.title();
//        }
//        else {
//            cout << "Sheet number is not valid. Try again." << endl;
//        }
//    }
//}
//
//string CopyFile()
//{
//    string originalFileName, NewFileName;
//
//    do
//    {
//        cout << "Provide original path name: ";
//        cin >> originalFileName;
//
//        if (!fs::exists(originalFileName))
//        {
//            cout << RED << "Error occured" << RESET << endl;
//        }
//        else
//        {
//            cout << GREEN << "Correct path" << RESET << endl;
//        }
//    } while (!fs::exists(originalFileName));
//
//    do
//    {
//        cout << "Provide new path: ";
//        cin >> NewFileName;
//
//        if (!fs::exists(NewFileName))
//        {
//            cout << RED << "Error occured" << RESET << endl;
//        }
//        else
//        {
//            cout << GREEN << "Correct path" << RESET << endl;
//        }
//    } while (!fs::exists(NewFileName));
//
//    string DestName = NewFileName + "/excel_new.xlsx";
//
//    fs::copy_file(originalFileName, DestName);
//
//    return DestName;
//
//}

int titleToNumber(string s)
{
    int r = 0;
    for (int i = 0; i < s.length(); i++)
    {
        r = r * 26 + s[i] - 64;
    }
    return r;
}

pair <int, int> signs_numbers_separately(string cell_input)
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

void get_cell(string first_cell, string second_cell)
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
//wchar_t* convertchararraytolpcwstr(const char* chararray)
//{
//    char text[] = "something";
//    wchar_t wtext[256];
//    mbstowcs(wtext, text, strlen(text) + 1);//Plus null
//    LPWSTR ptr = wtext;
//    return ptr;
//}
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
void Convert(string path, string sheet, string cl1, string cl2, string cd1, string cd2)
{
    string newFileName, originalFileName, newName;

    newName = path;

    xlnt::workbook wb;
    xlnt::worksheet ws;
    //Loading excel file
    try {

        wb.load(path);  // excel file path

        //ws = wb.active_sheet();//getting active worksheet from workbook
        ws = wb.sheet_by_title(sheet);//getting specified by user worksheet from workbook

    }
    catch (const xlnt::exception& e) {
        MessageBox(NULL, charToLPCWSTR(e.what()), L"Error", MB_ICONERROR);
        //cout << RED << "Processing failed " << RESET << e.what() << endl;//displaying an error message when trying to read a file
        return;
    }

    //LEGEND
    ///////////////////////////////////////////////////////////////////////////


    //cout << "Provide cells for legend" << endl;

    get_cell(cl1, cl2);
    double* min_cell_values = new double[(coordinates[2] - coordinates[0]) + 2];
    double* max_cell_values = new double[(coordinates[2] - coordinates[0]) + 2];
    //vector<int> min_cell_values;
    //vector<int> max_cell_values;

    vector<xlnt::rgb_color> color_table;

    xlnt::cell cell = ws.cell(1, 1);

    int col_counter;

    int index = 0;

    int length = coordinates[2] - coordinates[0];

    for (int i = coordinates[0]; i < (coordinates[2] + 1); i++)
    {
        col_counter = 1;
        for (int j = coordinates[1]; j < (coordinates[3] + 1); j++)
        {
            cell = ws.cell(i, j);
            switch (col_counter)
            {
            case 1:
            {
                color_table.push_back(get_cell_color_solid(cell_get(i, j, ws)));
                break;
            }
            case 2:
            {
                min_cell_values[index] = int(cell.value<int>());
                //cout << "min: " << min_cell_values[i] << endl;
                break;
            }
            case 3:
            {
                max_cell_values[index] = int(cell.value<int>());
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
    
    get_cell(cd1, cd2);

    color_fill_cell(cell_get(1, 1, ws), color_table[0]);

    for (int i = coordinates[1]; i < coordinates[3] + 1; i++)
    {
        for (int j = coordinates[0]; j < coordinates[2] + 1; j++)
        {

            //cout << to_string(cell_get(j, i, ws).value<int>()) << " ";

            for (int w = 0; w < length + 1; w++)
            {
                //cout << "w: " << w << endl;
                //cout << to_string(min_cell_values[w]) << endl;
                //cout << to_string(max_cell_values[w]) << endl;
                if (cell_get(j, i, ws).value<int>() > min_cell_values[w] && cell_get(j, i, ws).value<int>() < max_cell_values[w])
                {
                    color_fill_cell(cell_get(j, i, ws), color_table[w]);
                    break;
                }
            }

        }
        //cout << endl;
    }


    wb.save(newName);

}