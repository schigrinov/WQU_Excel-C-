using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace WorldQuant_Module3_CSA_SkeletonCode
{
    class Program
    {
        static Excel.Workbook workbook;
        static Excel.Application app;

        static void Main(string[] args)
        {
            app = new Excel.Application();
            app.Visible = true;
            try
            {
                workbook = app.Workbooks.Open("property_pricing.xlsx", ReadOnly: false);
            }
            catch
            {
                SetUp();
            }

            var input = "";
            while (input != "x")
            {
                PrintMenu();
                input = Console.ReadLine();
                try
                {
                    var option = int.Parse(input);
                    switch (option)
                    {
                        case 1:
                            try
                            {
                                Console.Write("Enter the size: ");
                                var size = float.Parse(Console.ReadLine());
                                Console.Write("Enter the suburb: ");
                                var suburb = Console.ReadLine();
                                Console.Write("Enter the city: ");
                                var city = Console.ReadLine();
                                Console.Write("Enter the market value: ");
                                var value = float.Parse(Console.ReadLine());

                                AddPropertyToWorksheet(size, suburb, city, value);
                            }
                            catch
                            {
                                Console.WriteLine("Error: couldn't parse input");
                            }
                            break;
                        case 2:
                            Console.WriteLine("Mean price: " + CalculateMean());
                            break;
                        case 3:
                            Console.WriteLine("Price variance: " + CalculateVariance());
                            break;
                        case 4:
                            Console.WriteLine("Minimum price: " + CalculateMinimum());
                            break;
                        case 5:
                            Console.WriteLine("Maximum price: " + CalculateMaximum());
                            break;
                        default:
                            break;
                    }
                } catch { }
            }

            // save before exiting
            workbook.Save();
            workbook.Close();
            app.Quit();
        }

        static void PrintMenu()
        {
            Console.WriteLine();
            Console.WriteLine("Select an option (1, 2, 3, 4, 5) " +
                              "or enter 'x' to quit...");
            Console.WriteLine("1: Add Property");
            Console.WriteLine("2: Calculate Mean");
            Console.WriteLine("3: Calculate Variance");
            Console.WriteLine("4: Calculate Minimum");
            Console.WriteLine("5: Calculate Maximum");
            Console.WriteLine();
        }

        static void SetUp()
        {
            app.Workbooks.Add();
            workbook = app.ActiveWorkbook;
            
            Excel.Worksheet currentSheet = workbook.Worksheets[1];
            currentSheet.Name = "Properties";
            currentSheet.Cells[1, 1] = "Size (in square feet)";
            currentSheet.Cells[1, 2] = "Suburb";
            currentSheet.Cells[1, 3] = "City";
            currentSheet.Cells[1, 4] = "Market Value";

            currentSheet.Cells[1, 7] = "Counter";
            currentSheet.Cells[2, 7] = 0;

            for (int i = 0; i < 7; i++)
            {
                currentSheet.Columns[i].AutoFit();
            }

            
            workbook.SaveAs("property_pricing.xlsx");
        }

        static void AddPropertyToWorksheet(float size, string suburb, string city, float value)
        {
            Excel.Worksheet currentSheet = workbook.Worksheets[1];
            var ind = currentSheet.Cells[2, 7].Value;
            ind += 1;
            currentSheet.Cells[2, 7] = ind;
            ind += 1;
            currentSheet.Cells[ind, 1] = size;
            currentSheet.Cells[ind, 2] = suburb;
            currentSheet.Cells[ind, 3] = city;
            currentSheet.Cells[ind, 4] = value;
        }

        static float CalculateMean()
        {
            Excel.Worksheet currentSheet = workbook.Worksheets[1];
            float total = 0f;
            int counter = (int)currentSheet.Cells[2, 7].Value;
            int lastRow = counter + 1;
            for (int i=2; i<= lastRow; i++)
            {
                total += (float)currentSheet.Cells[i, 4].Value;
            }
            return total/counter;
        }

        static float CalculateVariance()
        {
            float mean = CalculateMean();
            Excel.Worksheet currentSheet = workbook.Worksheets[1];
            float total = 0f;
            int counter = (int)currentSheet.Cells[2, 7].Value;
            int lastRow = counter + 1;
            for (int i = 2; i <= lastRow; i++)
            {
                var val = currentSheet.Cells[i, 4].Value;
                val = (val - mean) * (val - mean);
                total += val;
            }
            // Console.WriteLine("Calculating Sample vatiance");
            return total / (counter-1);
        }

        static float CalculateMinimum()
        {
            Excel.Worksheet currentSheet = workbook.Worksheets[1];
            int counter = (int)currentSheet.Cells[2, 7].Value;
            int lastRow = counter + 1;
            float min = float.MaxValue;
            for (int i = 2; i <= lastRow; i++)
            {
                float val = (float)currentSheet.Cells[i, 4].Value;
                if (val < min) min = val;
            }
            return min;
        }

        static float CalculateMaximum()
        {
            Excel.Worksheet currentSheet = workbook.Worksheets[1];
            int counter = (int)currentSheet.Cells[2, 7].Value;
            int lastRow = counter + 1;
            float max = float.MinValue;
            for (int i = 2; i <= lastRow; i++)
            {
                float val = (float)currentSheet.Cells[i, 4].Value;
                if (val > max) max = val;
            }
            return max;
        }
    }
}
