using ClosedXML.Excel;

using Newtonsoft.Json;

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DU_Industry_Price_Calculator
{
    public class Calculator
    {

        /*
         * 
         * Okay let's talk about what this is.
         * This should take in a json structure of the recipes, which we have
         * Then, we must input ore prices for all basic ores
         * 
         * It then should iterate through all recipes, calculating the Price To Make if you bought ores form the market, and processed it all with or without full talents
         * We then output all that information to a spreadsheet
         * */

        private Dictionary<string, Recipe> _recipes; // Global list of all our recipes, from the json
        private List<string> missing_recipes = new List<string>(); // Used to report any errors from missing recipes

        public Calculator()
        {
            // Step 1, load json
            var json = File.ReadAllText("recipes_object_version.json");
            _recipes = JsonConvert.DeserializeObject<Dictionary<string,Recipe>>(json);


            // Alright, so now iterate all recipes and calculate the Price To Make
            // And I guess, just list them all, except the stupid honeycombs if they're even included

            // I don't think they are.

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Without Talents");
                worksheet.Cell(1, 1).Value = "Name";
                worksheet.Cell(1, 2).Value = "Price To Make";
                worksheet.Cell(1, 4).Value = "Input Ore";
                worksheet.Cell(1, 5).Value = "Amount Required";
                worksheet.Cell(1, 6).Value = "Cost Per Ore";
                worksheet.Cell(1, 7).Value = "Total Cost";

                var talentSheet = workbook.Worksheets.Add("With Talents");
                talentSheet.Cell(1, 1).Value = "Name";
                talentSheet.Cell(1, 2).Value = "Price To Make";
                talentSheet.Cell(1, 4).Value = "Input Ore";
                talentSheet.Cell(1, 5).Value = "Amount Required";
                talentSheet.Cell(1, 6).Value = "Cost Per Ore";
                talentSheet.Cell(1, 7).Value = "Total Cost";

                var worksheetHeader = worksheet.Range(1, 1, 1, 7);
                var talentHeader = talentSheet.Range(1, 1, 1, 7);
                worksheetHeader.Style.Font.SetBold();
                talentHeader.Style.Font.SetBold();
                worksheetHeader.Style.Border.BottomBorder = XLBorderStyleValues.Medium;
                talentHeader.Style.Border.BottomBorder = XLBorderStyleValues.Medium;


                int row = 2; // Start after the header

                // The ores are first when we iterate, we will store their cell positions for reference in formulas later
                Dictionary<string, string> oreCells = new Dictionary<string, string>();

                int lastOreRow = -1; // Just to flag that it's unset.  TBH 0 would be fine but it feels wrong
                int numProcessed = 0;

                foreach (var recipekvp in _recipes)
                {
                    string name = recipekvp.Key; // Just for easier reference
                    var recipe = recipekvp.Value;

                    Dictionary<string, double> oreTotals;
                    try
                    {
                        oreTotals = getOreTotals(recipe);
                    }
                    catch (Exception)
                    {
                        // Had missing recipes, skip this recipe entirely, we'll list the problem later in the Missing Recipes tab
                        continue;
                    }
                    // If we were going to get an exception from this we already did, so, no need to slow it down with a try/catch this time
                    var talentTotals = getOreTotals(recipe, true);

                    if (recipe.Type == "Ore" || recipe.Name == "Hydrogen Pure" || recipe.Name == "Oxygen Pure")
                    {
                        worksheet.Cell(row, 1).Value = name;
                        worksheet.Cell(row, 2).Value = recipe.Price;
                        worksheet.Cell(row, 1).Style.Font.SetBold();
                        worksheet.Cell(row, 2).Style.Font.SetBold();

                        talentSheet.Cell(row, 1).Value = name;
                        talentSheet.Cell(row, 2).Value = recipe.Price;
                        talentSheet.Cell(row, 1).Style.Font.SetBold();
                        talentSheet.Cell(row, 2).Style.Font.SetBold();

                        oreCells[name] = "B" + row; // Store the cell location in a dictionary by name
                        worksheet.Cell(row, 2).Style.Fill.BackgroundColor = XLColor.AliceBlue; // Indicate that these cells are changeable
                        talentSheet.Cell(row, 2).Style.Fill.BackgroundColor = XLColor.AliceBlue;
                        row++;
                    }
                    else
                    {
                        if (lastOreRow == -1) // if this is the first non-ore, finish out the Ore section and add a few rows of space
                        {
                            worksheet.Range(row-1, 1, row-1, 7).Style.Border.BottomBorder = XLBorderStyleValues.Thin; // I magically know that 7 is the highest column... TODO: remove magic
                            talentSheet.Range(row - 1, 1, row - 1, 7).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                            lastOreRow = row;
                            row+=3;
                        }

                        worksheet.Cell(row, 1).Value = name;
                        // The cell at row,2 gets populated by a formula, after we iterate the ores and determine the range over which we should SUM
                        worksheet.Cell(row, 1).Style.Font.SetBold();
                        worksheet.Cell(row, 2).Style.Font.SetBold();

                        talentSheet.Cell(row, 1).Value = name;
                        talentSheet.Cell(row, 1).Style.Font.SetBold();
                        talentSheet.Cell(row, 2).Style.Font.SetBold();

                        int startRow = row; // Track the first row so we can put borders around the whole thing
                        row++;
                        // List the ores
                        foreach (var orekvp in oreTotals)
                        {
                            worksheet.Cell(row, 4).Value = orekvp.Key;
                            worksheet.Cell(row, 5).Value = orekvp.Value; // Amount Required, E
                            worksheet.Cell(row, 6).FormulaA1 = "=" + oreCells[orekvp.Key];
                            worksheet.Cell(row, 7).FormulaA1 = "=" + oreCells[orekvp.Key] + "*E" + row; // Amount Required * Price Per Unit

                            talentSheet.Cell(row, 4).Value = orekvp.Key;
                            talentSheet.Cell(row, 5).Value = talentTotals[orekvp.Key];
                            talentSheet.Cell(row, 6).FormulaA1 = "=" + oreCells[orekvp.Key];
                            talentSheet.Cell(row, 7).FormulaA1 = "=" + oreCells[orekvp.Key] + "*E" + row;
                            row++;
                        }
                        worksheet.Range(startRow, 1, row - 1, 7).Style.Border.OutsideBorder = XLBorderStyleValues.Double;
                        talentSheet.Range(startRow, 1, row - 1, 7).Style.Border.OutsideBorder = XLBorderStyleValues.Double;
                        if (numProcessed % 2 == 0) // Make the recipes have alternating colors for easier reading
                        {
                            worksheet.Range(startRow, 1, row - 1, 7).Style.Fill.BackgroundColor = XLColor.FloralWhite;
                            talentSheet.Range(startRow, 1, row - 1, 7).Style.Fill.BackgroundColor = XLColor.FloralWhite;
                        }
                        // Set the formulas for PriceToMake - the sum of all prices of the constituent ores
                        worksheet.Cell(startRow, 2).FormulaA1 = "=SUM(G" + (startRow + 1) + ":G" + (row - 1) + ")";
                        talentSheet.Cell(startRow, 2).FormulaA1 = "=SUM(G" + (startRow + 1) + ":G" + (row - 1) + ")";
                    }
                    numProcessed++; // This is just to help determine 'alternating' recipes for the background colors
                    Console.WriteLine($"Finished {recipe.Name}");
                }
                worksheet.ColumnsUsed().AdjustToContents();
                talentSheet.ColumnsUsed().AdjustToContents();

                if (missing_recipes.Count > 0)
                {
                    var missingSheet = workbook.Worksheets.Add("Missing Recipes");
                    int missingRow = 1;
                    foreach (var missing in missing_recipes)
                        missingSheet.Cell(missingRow++, 1).Value = missing;
                    missingSheet.Protect();
                }

                // Now, protect all the sheets
                worksheet.Protect();
                talentSheet.Protect();
                // And allow them to change column B up to lastOreRow
                worksheet.Range(2, 2, lastOreRow, 2).Style.Protection.SetLocked(false);
                talentSheet.Range(2, 2, lastOreRow, 2).Style.Protection.SetLocked(false);

                workbook.SaveAs("FinalData.xlsx");
            }
            Console.WriteLine("Workbook saved, done");
           
        }

        // This function was kindof a nightmare to get the logic right so I had to step through piece by piece
        private Dictionary<string, double> getOreTotals(Recipe recipe, bool assumeTalents = false)
        {
            var result = new Dictionary<string, double>(); // Example step through of Nitron.
            var numPerRecipe = recipe.OutputQuantity; // 100 Quantity
            var costMultiplier = 1d;
            if (assumeTalents && (recipe.Type == "Fuel" || recipe.Type == "Product" || recipe.Type == "Pure" || recipe.Type == "Scrap" || recipe.Type.Contains("Ammo")))
            {
                numPerRecipe *= 1.15f; // 15% more output with talents
                if (recipe.Type != "Fuel")
                    costMultiplier = 0.85d; // 15% less costs with talents
                else
                    costMultiplier = 0.75d; // Fuel gets -10% generic and -15% specific for costs, so only costs 75%
            }
            else if (assumeTalents && recipe.Type == "Intermediary Part")
            {
                numPerRecipe += 5; // Intermediate parts of all tiers can be given up to a flat +5 to outputs, but can't have reduced inputs 
            }

            foreach (KeyValuePair<string,double> input in recipe.Input)
            {
                if (!_recipes.ContainsKey(input.Key))
                {
                    missing_recipes.Add(input.Key); // This recipe and its parents will be aborted and not listed, and a separate tab will say what recipes were missing
                    throw new Exception("Stupid recipe missing");
                }
                else
                {
                    var inputRecipe = _recipes[input.Key]; // This is the recipe for one of our inputs, in our Nitron case it's the recipe for Pure Carbon or Silicon
                    var numNeeded = input.Value * costMultiplier / numPerRecipe; // This is how many Carbon Pure we need for 1 of our main recipe

                    if (inputRecipe.Type != "Ore")
                    {
                        // If it's not an ore, we recurse
                        // This should return the amount of Carbon Ore needed to make a single Carbon Pure
                        var innerResult = getOreTotals(inputRecipe, assumeTalents);
                        foreach (var kvp in innerResult) // And safely add them back in
                        {
                            var value = kvp.Value * numNeeded;
                            if (result.ContainsKey(kvp.Key))
                                result[kvp.Key] += value;
                            else
                                result[kvp.Key] = value;
                        }
                    }
                    else
                    {
                        // If it is an ore, we can just add it in
                        // numNeeded is appropriately scaled, that's how many of this ore we need after the discounts, and for only one of the main recipe
                        var value = numNeeded; 
                        if (result.ContainsKey(input.Key))
                            result[input.Key] += value;
                        else
                            result[input.Key] = value;
                    }
                }
            }
            if (recipe.Input.Count == 0 && recipe.Type == "Ore")
            {
                // It had no inputs and is a basic ore, make it just cost one of itself for display purposes
                result[recipe.Name] = 1;
            }
            return result;
        }
    }
}
