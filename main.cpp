#include <iostream>
#include <fstream>
#include <sstream>
#include <vector>
#include <OpenXLSX.hpp> // Include the OpenXLSX library

using namespace std;
using namespace OpenXLSX;

// Function to split a string based on a delimiter
vector<string> split(const string& str, char delimiter) {
    vector<string> tokens;
    string token;
    stringstream ss(str);
    while (getline(ss, token, delimiter)) {
        tokens.push_back(token);
    }
    return tokens;
}

int main() {
    string inputLine;
    cout << "Enter the text (use ',' for next cell and ';' for next row):" << endl;

    vector<vector<string>> table;

    // Read input until an empty line
    while (true) {
        getline(cin, inputLine);
        if (inputLine.empty()) break;

        // Split the line into rows based on ';'
        vector<string> rows = split(inputLine, ';');
        for (const string& row : rows) {
            table.push_back(split(row, ',')); // Split each row into cells by ','
        }
    }

    // Create an Excel workbook and worksheet
    XLDocument doc;
    doc.create("Output.xlsx");
    auto wks = doc.workbook().worksheet("Sheet1");

    // Write data to the Excel worksheet
    for (size_t i = 0; i < table.size(); ++i) {
        for (size_t j = 0; j < table[i].size(); ++j) {
            wks.cell(XLCellReference(i + 1, j + 1)).value() = table[i][j];
        }
    }

    // Save and close the workbook
    doc.save();
    doc.close();

    cout << "Excel file 'Output.xlsx' has been created successfully." << endl;
    return 0;
}
