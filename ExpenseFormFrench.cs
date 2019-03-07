using Microsoft.Office.InfoPath;
using System;
using System.Xml;
using System.Xml.XPath;

namespace MS_Expense_Form_French_V5
{
    public partial class FormCode
    {
        // Member variables are not supported in browser-enabled forms.
        // Instead, write and read these values from the FormState
        // dictionary using code such as the following:
        //
        // private object _memberVariable
        // {
        //     get
        //     {
        //         return FormState["_memberVariable"];
        //     }
        //     set
        //     {
        //         FormState["_memberVariable"] = value;
        //     }
        // }

        // NOTE: The following procedure is required by Microsoft InfoPath.
        // It can be modified using Microsoft InfoPath.
        public void InternalStartup()
        {
            ((ButtonEvent)EventManager.ControlEvents["CTRL140_10"]).Clicked += new ClickedEventHandler(CTRL140_10_Clicked);
        }
        public void CTRL140_10_Clicked(object sender, ClickedEventArgs e)
        {
            XPathNavigator nav2s = this.MainDataSource.CreateNavigator();
            //In the form, these fields are set from the rules of Main/myFields/ProvinceName, querying information from the Tax Calculation 
            //list in the site contents of mercury
            double gstRate = Convert.ToDouble(nav2s.SelectSingleNode("/my:myFields/my:GST", NamespaceManager).Value);
            double pstRate = Convert.ToDouble(nav2s.SelectSingleNode("/my:myFields/my:PST", NamespaceManager).Value);
            double gstRebate = Convert.ToDouble(nav2s.SelectSingleNode("/my:myFields/my:GSTRebate", NamespaceManager).Value);
            double pstRebate = Convert.ToDouble(nav2s.SelectSingleNode("/my:myFields/my:PSTRebate", NamespaceManager).Value);
            //string isThereTax = nav2s.SelectSingleNode("/my:myFields/my:isThereTax", NamespaceManager).Value;
            //initialization of tax variables and set to default value of 0
            double subtotal, valueBeforeTax, gst, pst, totalExpense, totalTaxableItems, totalNonTaxableItems;
            string isThereTax;
            subtotal = valueBeforeTax = gst = pst = totalExpense = totalTaxableItems = totalNonTaxableItems = 0;
            subtotal = Convert.ToDouble(nav2s.SelectSingleNode("/my:myFields/my:SubTotal", NamespaceManager).Value);
            //tax calculations
            /*if (isThereTax == "true") 
            {
                total = Convert.ToDouble(nav2s.SelectSingleNode("/my:myFields/my:SubTotal", NamespaceManager).Value);
                valueBeforeTax = total / (1 + gstRate + pstRate);
                gst = valueBeforeTax * gstRate * (1 - gstRebate);
                pst = valueBeforeTax * pstRate * (1 - pstRebate);
                totalExpense = valueBeforeTax + gst + pst;
            }*/
            string pathToTable = "/my:myFields/my:LineItems";
            //points and iterator node to the table of items to iterate through the rows
            XPathNodeIterator iterate = nav2s.Select(pathToTable, NamespaceManager);
            iterate.MoveNext(); //We do this as the first row is blank and not filled with fields

            //
            int numOfRows = iterate.Count;
            string[] arrayOfItems = new string[numOfRows];
            for (int i = 1; i < iterate.Count + 1; i++)
            {
                arrayOfItems[i - 1] = nav2s.SelectSingleNode(pathToTable + "[" + i.ToString() + "]" + "/my:InputAmount", NamespaceManager).Value;
                isThereTax = nav2s.SelectSingleNode(pathToTable + "[" + (i).ToString() + "]" + "/my:Taxable", NamespaceManager).Value; //this checks if that line item is taxable
                double value = Convert.ToDouble(arrayOfItems[i - 1]);
                if (isThereTax == "true")
                {
                    totalTaxableItems += value;
                }
                else totalNonTaxableItems += value;
            }
            if (totalNonTaxableItems != subtotal)//i.e. means there are tax items
            {
                valueBeforeTax = totalTaxableItems / (1 + gstRate + pstRate);
                gst = valueBeforeTax * gstRate * (1 - gstRebate);
                pst = valueBeforeTax * pstRate * (1 - pstRebate);
                totalExpense = valueBeforeTax + gst + pst;
            }
            for (int i = 1; i < iterate.Count + 1; i++)//this loops through each row item
            {
                //Grabs the inputed amount from the form per row item
                arrayOfItems[i - 1] = nav2s.SelectSingleNode(pathToTable + "[" + i.ToString() + "]" + "/my:InputAmount", NamespaceManager).Value;
                isThereTax = nav2s.SelectSingleNode(pathToTable + "[" + i.ToString() + "]" + "/my:Taxable", NamespaceManager).Value; //this checks if that line item is taxable
                double value = Convert.ToDouble(arrayOfItems[i - 1]);
                if (isThereTax == "false")//if no tax then inputted amount = item amount
                {
                    nav2s.SelectSingleNode(pathToTable + "[" + (i).ToString() + "]" + "/my:ItemAmount", NamespaceManager).SetValue(value.ToString());
                }
                if (isThereTax == "true")//changes item amount based on the GST and PST calculations
                {
                    double finalValue = value / totalTaxableItems * totalExpense;
                    double final = Math.Round(finalValue, 2, MidpointRounding.AwayFromZero);
                    nav2s.SelectSingleNode(pathToTable + "[" + (i).ToString() + "]" + "/my:ItemAmount", NamespaceManager).SetValue(final.ToString());
                }

            }

            //Now Fill in the TaxTable seen during Finance Review and Entry PageView

            string path = "/my:myFields/my:TaxTable/my:TaxItems";
            // Create a navigator object to point to the root of the data source of the VP form
            XPathNavigator tableNav;
            tableNav = this.MainDataSource.CreateNavigator();
            tableNav.SelectSingleNode(path, NamespaceManager);
            // Navigator object selects a group of nodes, specified by the path
            // Get the Iterator object to point to the table (group of nodes) specified by the path (currently not pointed to the first node in the set of nodes)
            // Object will iterate over the group of nodes 
            XPathNodeIterator rows = tableNav.Select(path, NamespaceManager);

            XPathNavigator first = tableNav.SelectSingleNode("//my:myFields/my:TaxTable/my:TaxItems[2]", NamespaceManager);
            XPathNavigator last;
            // If more than one row exists, means the Calculae Tax button has been selected before therefore we need to delete the current rows and add new ones

            if (rows.Count > 1)
            {
                last = tableNav.SelectSingleNode("//my:myFields/my:TaxTable/my:TaxItems[2]", NamespaceManager);
                // Delete the current node (first) to the node specified using "last"
                first.DeleteRange(last);
            }
            rows.MoveNext();

            //Fill in the PSTRebate Row

            // rows.Current causes the XPathNavigator object to point to the current node (row) and selects a node (fields) within that node/row
            rows.Current.SelectSingleNode("/my:myFields/my:TaxTable/my:TaxItems/my:taxName", NamespaceManager).SetValue("PST Rebate");
            string pstGL = nav2s.SelectSingleNode("/my:myFields/my:PSTGLCode", NamespaceManager).Value;//grabs the PSTGL code from the Tax Calculation List
            rows.Current.SelectSingleNode("/my:myFields/my:TaxTable/my:TaxItems/my:GLCodeTax", NamespaceManager).SetValue(pstGL);
            if (totalNonTaxableItems != subtotal)
            {
                double rebatePST = Math.Round(valueBeforeTax * pstRate * pstRebate, 2, MidpointRounding.AwayFromZero);
                nav2s.SelectSingleNode("/my:myFields/my:PSTRebateTotal", NamespaceManager).SetValue(rebatePST.ToString());
                rows.Current.SelectSingleNode("/my:myFields/my:TaxTable/my:TaxItems/my:TaxAmount", NamespaceManager).SetValue(rebatePST.ToString());
            }
            else
            {
                nav2s.SelectSingleNode("/my:myFields/my:PSTRebateTotal", NamespaceManager).SetValue("0");
                rows.Current.SelectSingleNode("/my:myFields/my:TaxTable/my:TaxItems/my:TaxAmount", NamespaceManager).SetValue("0");
            }



            //Play around with this section
            // Have the iterator object point to the beginning of the table again by once again assigning a collection of nodes to the iterator
            rows = tableNav.Select(path, NamespaceManager);
            rows.MoveNext();
            // Add another row by cloning the PST tax row and adding the clones row to the bottom of the table and overriding the original row with a GST row
            rows.Current.Clone();
            rows.Current.InsertAfter(rows.Current);


            //repeat same process for GST Table

            rows.Current.SelectSingleNode("/my:myFields/my:TaxTable/my:TaxItems/my:taxName", NamespaceManager).SetValue("GST Rebate");
            string gstGL = nav2s.SelectSingleNode("/my:myFields/my:GSTGLCode", NamespaceManager).Value;
            rows.Current.SelectSingleNode("/my:myFields/my:TaxTable/my:TaxItems/my:GLCodeTax", NamespaceManager).SetValue(gstGL);

            //if (isThereTax == "true")
            //{
            double rebateGST = Math.Round(valueBeforeTax * gstRate * gstRebate, 2, MidpointRounding.AwayFromZero);
            rows.Current.SelectSingleNode("/my:myFields/my:TaxTable/my:TaxItems/my:TaxAmount", NamespaceManager).SetValue(rebateGST.ToString());
            nav2s.SelectSingleNode("/my:myFields/my:GSTRebateTotal", NamespaceManager).SetValue(rebateGST.ToString());
            //}
            //else
            //{
            //    rows.Current.SelectSingleNode("/my:myFields/my:TaxTable/my:TaxItems/my:TaxAmount", NamespaceManager).SetValue("0");
            //    nav2s.SelectSingleNode("/my:myFields/my:GSTRebateTotal", NamespaceManager).SetValue("0");
            //}





            // Adjust the adjusted amount in the first form to ensure that the sum of the rebates + the adjusted amounts is equal to total of the inputted values
            if (totalNonTaxableItems != subtotal) adjustFirstItem();



        }

        public void adjustFirstItem()
        {
            // The purpose of this function is to ensure that the sum total of the rebates + the adjusted amounts is equal to the total of the inputted amounts
            // Set difference variable and paths to the current calculated total of the input values (SubTotal), current calculated GST rebate, current calculated PST rebate, and the current sum of the adjusted values
            double difference;
            XPathNavigator nav = this.MainDataSource.CreateNavigator();
            string pathToInvoiceTotal = "/my:myFields/my:SubTotal";
            string pathToGSTRebate = "/my:myFields/my:GSTRebateTotal";
            string pathToPSTRebate = "/my:myFields/my:PSTRebateTotal";
            string pathToAdjustedAmountSum = "/my:myFields/my:currentSum";

            // Extract the current Invoice Total value, GST Rebate Total, PST Rebate total, and the total value of the adjusted amounts using the XPathNavigator
            string currentInvoiceTotal = nav.SelectSingleNode(pathToInvoiceTotal, NamespaceManager).Value;
            string GSTRebate = nav.SelectSingleNode(pathToGSTRebate, NamespaceManager).Value;
            string PSTRebate = nav.SelectSingleNode(pathToPSTRebate, NamespaceManager).Value;
            string adjustedAmountTotal = nav.SelectSingleNode(pathToAdjustedAmountSum, NamespaceManager).Value;

            // Convert all string values (from above) to doubles and round them to two decimal places as InfoPath stores all fields as strings
            // Infopath displays 2 decimals however, numbers actually contain many more therefore rounding is mandatory and MidpointRounding function is defaulted to round down at midpoint values (i.e 1.5 or 65)
            double Invoice = Math.Round(Convert.ToDouble(currentInvoiceTotal), 2, MidpointRounding.AwayFromZero);
            double gstRebate = Math.Round(Convert.ToDouble(GSTRebate), 2, MidpointRounding.AwayFromZero);
            double pstRebate = Math.Round(Convert.ToDouble(PSTRebate), 2, MidpointRounding.AwayFromZero);
            double adjustedSum = Convert.ToDouble(adjustedAmountTotal);

            // Calculate the current total from summing the gstRebate, pstRebate, and adjusted values sum. We want this sum to eqaul the sum of the inputted line items
            double currentSum = gstRebate + pstRebate + adjustedSum;

            // Find the adjusted value currently in the last line item 
            XPathNodeIterator iterator = nav.Select("/my:myFields/my:LineItems", NamespaceManager);

            for (int counter = 0; counter < (iterator.Count + 1); counter++)
            {
                iterator.MoveNext();
            }
            // Convert the adjusted value in the last item to a double
            string lastItem = iterator.Current.SelectSingleNode("/my:myFields/my:LineItems/my:ItemAmount", NamespaceManager).Value;
            double lastItemVal = Convert.ToDouble(lastItem);


            // Compare the current input total and total sum of the adjusted values + rebates
            // If the two are equal, do nothing
            if (currentSum == Invoice) { }
            else
            {
                // If the two are not equal, find the difference and add it or subtract (add when the total of the adjusted + rebate is less than the total of the input values, subtract when the total of the adjusted + rebate is larger)
                if (currentSum > Invoice)
                {
                    // Find the difference and subtract it from the line item adjusted value shown 
                    difference = Math.Round(currentSum - Invoice, 2, MidpointRounding.AwayFromZero);
                    lastItemVal = Math.Round(lastItemVal - difference, 2, MidpointRounding.AwayFromZero);


                }
                else
                {
                    // Find the difference and add it to the line item adjusted value shown 
                    difference = Math.Round(Invoice - currentSum, 2, MidpointRounding.AwayFromZero);
                    lastItemVal = Math.Round(lastItemVal + difference, 2, MidpointRounding.AwayFromZero);

                }

            }

            // Put the new value of the adjusted line item into the table 
            iterator = nav.Select("/my:myFields/my:LineItems", NamespaceManager);

            for (int counter = 0; counter < (iterator.Count + 1); counter++)
            {
                iterator.MoveNext();
            }
            iterator.Current.SelectSingleNode("/my:myFields/my:LineItems/my:ItemAmount", NamespaceManager).SetValue(lastItemVal.ToString());
        }
    }
}
