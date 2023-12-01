#include <string>
#include <iostream>
#include <fstream>
#include <filesystem>

#include <windows.h>
#include "globals.h"
#include <string>
#include <codecvt>
#include <locale>

#include "libxl.h"
#include <map>
//define colors
#define RED "\033[31m"    
#define GREEN "\033[32m" 
#define RESET "\033[0m"

using namespace libxl;

//global variables
/*std::string announcement, color;
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
*/

void parseCellReference(const std::string& ref, int& row, int& col) {
    col = 0;
    int i = 0;

    // Process alphabetic part for column (e.g., "A", "B", ..., "AA", ...)
    while (i < ref.length() && std::isalpha(ref[i])) {
        col = col * 26 + (ref[i] - 'A' + 1);
        i++;
    }

    col--; // Adjust column index to be zero-based

    // Process numeric part for row
    row = 0;
    while (i < ref.length() && std::isdigit(ref[i])) {
        row = row * 10 + (ref[i] - '0');
        i++;
    }

    row--; // Adjust row index to be zero-based
}
wchar_t* StringToWChar_t(const std::string& str) {
    int len = MultiByteToWideChar(CP_UTF8, 0, str.c_str(), -1, NULL, 0);
    wchar_t* wstr = new wchar_t[len];
    MultiByteToWideChar(CP_UTF8, 0, str.c_str(), -1, wstr, len);
    return wstr;
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
using namespace std;
void Convert(string path, string sheet, string cl1, string cl2, string cd1, string cd2)
{
    string newFileName, originalFileName, newName;

    newName = path;
    Book* book = xlCreateXMLBook();
    // Only threshhold is being used
    std::map<float, Format*> colorMap;
    if (book) {
        if (book->load(StringToWChar_t(path)))
        {
			Sheet* sheet = book->getSheet(8);
            if (sheet)
            {
                //Load legend (chyba tak to jest po angielsku)
                int startRow, startCol, endRow, endCol;
                parseCellReference(cl1, startRow, startCol);
                parseCellReference(cl2, endRow, endCol);
                
                for (int i = startCol; i <= endCol; i++)
                {
                    //Color
                    Format* format = sheet->cellFormat(startRow, i);
                    //Max value
                    CellType cellType = sheet->cellType(startRow, i+2);
                    if (cellType == CELLTYPE_NUMBER)
                    {
						float value = sheet->readNum(startRow, i+2);
						colorMap[value] = format;
					}
                    else
                    {
                        //Cos sie zjebalo
                    }
				}
			}
		}
    }

    /*
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
                min_cell_values[index] = double(cell.value<double>());
                //cout << "min: " << min_cell_values[i] << endl;
                break;
            }
            case 3:
            {
                max_cell_values[index] = double(cell.value<double>());
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
            xlnt::cell cell = cell_get(j, i, ws);
            //cout << to_string(cell_get(j, i, ws).value<int>()) << " ";
            if (cell.data_type() != xlnt::cell_type::number)
            {
				continue;
            }
            else {
                for (int w = 0; w < length + 1; w++)
                {
                    auto val = cell_get(j, i, ws).value<double>();

                    if (val > min_cell_values[w] && val < max_cell_values[w])
                    {
                        color_fill_cell(cell_get(j, i, ws), color_table[w]);
                        break;
                    }
                }
            }
        }
        //cout << endl;
    }


    wb.save(newName);
    */
}