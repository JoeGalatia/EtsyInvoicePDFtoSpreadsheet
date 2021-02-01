#include "NumberDuck.h"
#include <iostream>
#include <string>
#include <vector>

using namespace NumberDuck;

struct address {
	std::string streetName;
	int houseNumber;
	std::string country;
	std::string street;
	std::string state;
	std::string zipCode;
	std::string country;
	std::string compilation;
};

struct squishMellow {
	std::string pillowName;
	int pillowsOrdered;
};

struct invoiceData {
	unsigned int transactionNumber;
	std::string customerName;
	address customerAddress;
	squishMellow product;
	std::string orderDate;
	std::string paymentMethod;
	bool deliveryConfirmed;
	double money_recieved;
};

int main() {

	//Get amount of invoices // amount of pdfs in a folder // put number into total Orders
	int totalOrders = 157;
	std::vector<invoiceData> invoices;

	//Get Data from PDF // put data into invoices vector



	//Output Data to .xsl
	{
		Workbook workbook("");
		Worksheet* pWorksheet = workbook.GetWorksheetByIndex(0);

	//Titles
		//Cell* pCell = pWorksheet->GetCellByAddress("A1");
		//pCell->SetString("Customer Name");
		pWorksheet->GetCell(0, 0)->SetString("Transaction #");
		pWorksheet->GetCell(2, 0)->SetString("Customer Name");
		pWorksheet->GetCell(4, 0)->SetString("Address");
		pWorksheet->GetCell(6, 0)->SetString("Squishmellow Type");
		pWorksheet->GetCell(8, 0)->SetString("Date Ordered");
		pWorksheet->GetCell(10, 0)->SetString("Payment Method");
		pWorksheet->GetCell(12, 0)->SetString("Delivery Confirmed");
		pWorksheet->GetCell(14, 0)->SetString("CA$ Revenue");
		pWorksheet->GetCell(16, 0)->SetString("CA$ Expenses");
		pWorksheet->GetCell(18, 0)->SetString("CA$ Profit");

	//Format
	//Transaction # ',' Customer Name ',' Address ',' Squishmellow recieved ie Pink axolotl ',' Date Ordered ',' Payment Method ',' "Delivery Confirmed: "true/false ',' CA $payment recieved for pillow ',' CA $payment made for pillow ',' CA $Profit
		
		for (int i = 1; i <= totalOrders; i++)
		{
			pWorksheet->GetCell(0, i)->SetFloat(i);
			pWorksheet->GetCell(2, i)->SetString(invoices[i].customerName.c_str());
			pWorksheet->GetCell(4, i)->SetString(invoices[i].customerAddress.compilation.c_str());
			pWorksheet->GetCell(6, i)->SetString(invoices[i].product.pillowName.c_str());
			pWorksheet->GetCell(8, i)->SetString(invoices[i].orderDate.c_str());
			pWorksheet->GetCell(10, i)->SetString(invoices[i].paymentMethod.c_str());
			pWorksheet->GetCell(12, i)->SetBoolean(invoices[i].deliveryConfirmed);
			pWorksheet->GetCell(14, i)->SetFloat(invoices[i].money_recieved);
		}

		workbook.Save("EtsyInvoices.xls");

		/*debugging
		Workbook* pWorkbookIn = new Workbook("");
		if (pWorkbookIn->Load("EtsyInvoices.xls"))
		{
			Worksheet* pWorksheetIn = pWorkbookIn->GetWorksheetByIndex(0);
			Cell* pCellIn = pWorksheetIn->GetCell(0, 0);
			printf("Cell Contents: %s\n", pCellIn->GetString());
		}*/
	}

	return 0;
}