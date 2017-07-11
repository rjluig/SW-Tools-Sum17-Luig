/***********************************************************************
The program below is taken the document "A brief introduction 
to C++ and Interfacing with Excel" by Andrew L. Hazel.  Very few
modifications were made to the code.  The complete document can be
found at http://www.maths.manchester.ac.uk/~ahazel/EXCEL_C++.pdf
Additional steps are needed to use Visual Studio Express Edition.
See footnote on page 58 of text.

***********************************************************************/

// Include standard libraries

#include<iostream>
#include<cmath>

// Import necessary Excel libraries.  Adjust the paths as necessary.

#import "C:\Program Files (x86)\Common Files\microsoft shared\OFFICE15\MSO.DLL" \
	rename("DocumentProperties", "DocumentPropertiesXL") \
	rename("RGB", "RBGXL")

#import  "C:\Program Files (x86)\Common Files\microsoft shared\VBA\VBA6\VBE6EXT.OLB"

#import  "C:\Program Files (x86)\Microsoft Office\Office15\EXCEL.EXE" \
	rename("DialogBox", "DialogBoxXL") \
	rename("RGB", "RBGXL") \
	rename("DocumentProperties", "DocumentPropertiesXL") \
	rename("ReplaceText", "ReplaceTextXL")	 \
	rename("CopyFile", "CopyFileXL") \
	exclude("IFont", "IPicture") no_dual_interfaces

using namespace std;

// Simple function to graph as an example of using Excel charting
// tools from within C++

double f(const double &x) { return (sin(x)*exp(-x)); }

int main()
{
	//Surround the entire interfacing code with a try block
	try
	{
		//Initialise the COM interface
		CoInitialize(NULL);
		//Define a pointer to the Excel	application
		Excel::_ApplicationPtr xl;
		//Start	one instance of Excel
		xl.CreateInstance(L"Excel.Application");
		//Make the Excel application visible
		xl->Visible = true;
		//Add a(new)workbook
		xl->Workbooks->Add(Excel::xlWorksheet);
		//Get a pointer	to the active worksheet
		Excel::_WorksheetPtr pSheet = xl->ActiveSheet;
		//Set the name of the sheet
		pSheet->Name = "Chart Data";
		//Get a pointer to the cells on the active worksheet
		Excel::RangePtr pRange = pSheet->Cells;
		//Define the number of plot points 
		unsigned Nplot = 100;
		//Set the lower and upper limits for x
		double x_low = 0.0, x_high = 20.0;
		//Calculate the size of the(uniform) x interval
		//Note a cast to a double here
		double h = (x_high - x_low) / (double)Nplot;
		//Create two columns of data in the worksheet
		//We put labels at the top of each column to say what it contains
		pRange->Item[1][1] = "x";
		pRange->Item[1][2] = "f(x)";
		//Now we fill in the rest of the actual data by
		//using a single for loop
		for (unsigned i = 0; i<Nplot; i++)
		{
			//Calculate the value of x (equally - spaced over the range)
			double x = x_low + i*h;
			//The first column is our equally - spaced x values
			pRange->Item[i + 2][1] = x;
			//The second column is f(x)
			pRange->Item[i + 2][2] = f(x);
		}
		//The sheet "Chart Data" now contains all the data
		//required to generate the chart
		//In order to use the Excel Chart Wizard,
		//we must convert the data into Range Objects
		//Set a pointer to the first cell containing our data
		Excel::RangePtr pBeginRange = pRange->Item[1][1];
		//Set a pointer to the last cell containing our data
		Excel::RangePtr pEndRange = pRange->Item[Nplot + 1][2];
		//Make a "composite" range of the pointers to the start
		//and end of our data
		//Note the casts to pointers to Excel Ranges
		Excel::RangePtr pTotalRange = pSheet->Range[(Excel::Range*)pBeginRange][(Excel::Range*)pEndRange];
		// Create the chart as a separate chart item in the workbook
		Excel::_ChartPtr pChart = xl->ActiveWorkbook->Charts->Add();
		//Use the ChartWizard to draw the chart.
		//The arguments to the chart wizard are
		//Source: the data range,
		//Gallery: the chart type,
		//Format: a chart format (number 1 - 10),
		//PlotBy: whether the data is stored in columns or rows,
		//CategoryLabels: an index for the number of columns
		//containing category (x) labels
		//(because our first column of data represents
		//the x values, we must set this value to 1)
		//SeriesLabels: an index for the number of rows containing
		//series (y) labels
		//(our first row contains y labels,
		//so we set this to 1)
		//HasLegend: boolean set to true to include a legend
		//Title: the title of the chart
		//CategoryTitle: the x - axis title
		//ValueTitle: the y - axis title
		pChart->ChartWizard((Excel::Range*)pTotalRange,
			(long)Excel::xlXYScatter,
			6L, (long)Excel::xlColumns, 1L, 1L, true,
			"My Graph", "x", "f(x)");
		//Give the chart sheet a name
		pChart->Name = "My Data Plot";
	}
	//If there has been an error, say so
	catch (_com_error &error)
	{
		cout << "COM ERROR" << endl;
	}
	//Finally Uninitialise the COM interface
	CoUninitialize();
	//Finish the C++ program
	return 0;
}
