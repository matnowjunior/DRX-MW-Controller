#include <string>
#include <iostream>
#include <xlnt/xlnt.hpp>
#include <fstream>

#define RED "\033[31m"    
#define GREEN "\033[32m" 
#define RESET "\033[0m"

using namespace std;

string announcement; //global variable

string CopyFile(string originalFileName, string newFileName)
{
    ifstream source(originalFileName, ios::binary); //opening file in binary mode
    ofstream dest(newFileName, ios::binary); //saving file to new one   

    if (source && dest) { //checking if both streams were opened correctly
        dest << source.rdbuf(); //coping the content of the source file to the target file
        source.close();//closing soource file
        dest.close();//closing target file
        announcement = "Excel file has been sucessfully copied";
    }
    else {
        announcement = "An error occurred while copying and renaming the file.";
    }
    return newFileName;
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

int main()
{
    int i, j;
    string newFileName, originalFileName, cell_input1, cell_input2;


    cout << "Provide the excel original file path: " << endl;
    cin >> originalFileName;
    cout << "Provide the excel new file path" << endl;
    cin >> newFileName;

    //Opening Excel workbook and worksheet
    xlnt::workbook wb;
    xlnt::worksheet ws;

    string newName = CopyFile(originalFileName, newFileName);

    //Loading excel file
    try {
        wb.load(newName);  // excel file path
        ws = wb.active_sheet();//getting active worksheet from workbook

    }
    catch (const xlnt::exception& e) {
        cout << RED << "Processing failed " << RESET << e.what() << std::endl;//displaying an error message when trying to read a file
        return 1;
    }



    cout << "Podaj pierwsza komorke:";
    cin >> cell_input1;


    int col_number1 = signs_numbers_separately(cell_input1).first;
    int row_number1 = signs_numbers_separately(cell_input1).second;

    cout << "Podaj druga komorke:";
    cin >> cell_input2;

    int col_number2 = signs_numbers_separately(cell_input2).first;
    int row_number2 = signs_numbers_separately(cell_input2).second;


    //looping through (before) specified range
    for (i = row_number1; i < row_number2 + 1; i++)
    {

        for (j = col_number1; j < col_number2 + 1; j++)
        {
            //setting a new cell to the entered values(col and row)
            xlnt::cell cell = ws.cell(i, j);
            float x = cell.value<float>();


            if (cell.has_value())
            {
                cout << cell.value<int>();

                if (x >= 0 && x < 30)
                    cell.fill(xlnt::fill::solid(xlnt::rgb_color(242, 34, 15)));
                else if (x >= 30 && x < 50)
                    cell.fill(xlnt::fill::solid(xlnt::rgb_color(242, 160, 7)));
                else if (x >= 50 && x < 75)
                    cell.fill(xlnt::fill::solid(xlnt::rgb_color(5, 242, 108)));
                else if (x >= 75 && x < 90)
                    cell.fill(xlnt::fill::solid(xlnt::rgb_color(5, 151, 242)));
                else if (x >= 90 && x < 100)
                    cell.fill(xlnt::fill::solid(xlnt::rgb_color(178, 37, 217)));
                cout << " ";
            }
            else
            {
                cout << "0";
                cout << " ";
            }

            //cell.clear_style();


        }
        cout << endl;
    }



    wb.save(newName);
    cout << announcement << endl;

    cout << GREEN << "Processing succeed :)" << RESET;
}