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
         * You can then select items one at a time, check their market sell price, and see if they're profitable
         * 
         * We can also export market prices from Hyperion, though idk how accurate they are
         * */

        private Dictionary<string, Recipe> _recipes;
        private List<string> missing_recipes = new List<string>();

        public Calculator()
        {
            // Step 1, load json
            var json = File.ReadAllText("recipes_object_version.json");
            _recipes = JsonConvert.DeserializeObject<Dictionary<string,Recipe>>(json);


            // Alright, so now iterate all recipes and calculate the Price To Make
            // And I guess, just list them all, except the stupid honeycombs if they're even included


            // Right so, turns out putting all these on a winform is, obviously, too much.
            // We'll put them in an excel sheet in the same way

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


                int row = 2;

                // Let's list all the Types just for posterity too
                List<string> types = new List<string>();

                // The ores are first, we will store their cell positions
                Dictionary<string, string> oreCells = new Dictionary<string, string>();

                int lastOreRow = -1;
                int numProcessed = 0;

                foreach (var kvp in _recipes)
                {
                    string key = kvp.Key;
                    var recipe = kvp.Value;

                    if (!types.Contains(recipe.Type))
                        types.Add(recipe.Type);

                    Dictionary<string, double> oreTotals;
                    try
                    {
                        oreTotals = getOreTotals(recipe);
                    }
                    catch (Exception)
                    {
                        // Had some issues with missing recipes, don't list it
                        continue;
                    }
                    // If we were going to get an exception from this we already did, so, no need to slow it down with a try/catch this time
                    var talentTotals = getOreTotals(recipe, true);

                    if (recipe.Type == "Ore" || recipe.Name == "Hydrogen Pure" || recipe.Name == "Oxygen Pure")
                    {
                        worksheet.Cell(row, 1).Value = kvp.Key;
                        worksheet.Cell(row, 2).Value = kvp.Value.Price;
                        worksheet.Cell(row, 1).Style.Font.SetBold();
                        worksheet.Cell(row, 2).Style.Font.SetBold();

                        talentSheet.Cell(row, 1).Value = kvp.Key;
                        talentSheet.Cell(row, 2).Value = kvp.Value.Price;
                        talentSheet.Cell(row, 1).Style.Font.SetBold();
                        talentSheet.Cell(row, 2).Style.Font.SetBold();

                        oreCells[kvp.Key] = "B" + row;
                        worksheet.Cell(row, 2).Style.Fill.BackgroundColor = XLColor.AliceBlue; // Indicate that these cells are changeable
                        talentSheet.Cell(row, 2).Style.Fill.BackgroundColor = XLColor.AliceBlue;
                        row++;
                    }
                    else
                    {
                        if (lastOreRow == -1) // if this is the first non-ore, finish out the Ore section and add a few rows of space
                        {
                            worksheet.Range(row-1, 1, row-1, 7).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                            talentSheet.Range(row - 1, 1, row - 1, 7).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                            lastOreRow = row;
                            row+=3;
                        }

                        worksheet.Cell(row, 1).Value = kvp.Key;
                        // We no longer pre-calculate prices, the excel sheet can do that - formulas get added after we iterate the ores
                        //worksheet.Cell(row, 2).Value = kvp.Value.PriceToMake;
                        worksheet.Cell(row, 1).Style.Font.SetBold();
                        worksheet.Cell(row, 2).Style.Font.SetBold();

                        talentSheet.Cell(row, 1).Value = kvp.Key;
                        //talentSheet.Cell(row, 2).Value = kvp.Value.PriceToMakeWithTalents;
                        talentSheet.Cell(row, 1).Style.Font.SetBold();
                        talentSheet.Cell(row, 2).Style.Font.SetBold();

                        int startRow = row;
                        row++;
                        // List the ores
                        foreach (var orekvp in oreTotals)
                        {
                            worksheet.Cell(row, 4).Value = orekvp.Key;
                            worksheet.Cell(row, 5).Value = orekvp.Value; // Amount Required, E
                            worksheet.Cell(row, 6).FormulaA1 = "=" + oreCells[orekvp.Key];
                            worksheet.Cell(row, 7).FormulaA1 = "=" + oreCells[orekvp.Key] + "*E" + row;

                            talentSheet.Cell(row, 4).Value = orekvp.Key;
                            talentSheet.Cell(row, 5).Value = talentTotals[orekvp.Key];
                            talentSheet.Cell(row, 6).FormulaA1 = "=" + oreCells[orekvp.Key];
                            talentSheet.Cell(row, 7).FormulaA1 = "=" + oreCells[orekvp.Key] + "*E" + row;
                            row++;
                        }
                        worksheet.Range(startRow, 1, row - 1, 7).Style.Border.OutsideBorder = XLBorderStyleValues.Double;
                        talentSheet.Range(startRow, 1, row - 1, 7).Style.Border.OutsideBorder = XLBorderStyleValues.Double;
                        if (numProcessed % 2 == 0)
                        {
                            worksheet.Range(startRow, 1, row - 1, 7).Style.Fill.BackgroundColor = XLColor.FloralWhite;
                            talentSheet.Range(startRow, 1, row - 1, 7).Style.Fill.BackgroundColor = XLColor.FloralWhite;
                        }
                        // Replace the Price To Make with the calculated values now that we know what range to sum over
                        worksheet.Cell(startRow, 2).FormulaA1 = "=SUM(G" + (startRow + 1) + ":G" + (row - 1) + ")";
                        talentSheet.Cell(startRow, 2).FormulaA1 = "=SUM(G" + (startRow + 1) + ":G" + (row - 1) + ")";
                    }
                    numProcessed++;
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

        private Dictionary<string, double> getOreTotals(Recipe recipe, bool assumeTalents = false)
        {
            var result = new Dictionary<string, double>(); // Run through of Nitron.
            var numPerRecipe = recipe.OutputQuantity; // 100 Quantity
            var costMultiplier = 1d;
            if (assumeTalents && (recipe.Type == "Fuel" || recipe.Type == "Product" || recipe.Type == "Pure" || recipe.Type == "Scrap" || recipe.Type.Contains("Ammo")))
            {
                numPerRecipe *= 1.15f; // 15% more output with talents
                if (recipe.Type != "Fuel")
                    costMultiplier = 0.85d; // 15% less costs with talents
                else
                    costMultiplier = 0.75d; // Fuel gets -10% generic and -15% specific for costs
            }
            else if (assumeTalents && recipe.Type == "Intermediary Part")
            {
                numPerRecipe += 5; // Intermediate parts of all tiers can be given up to a flat +5 to outputs, but can't have reduced inputs 
            }

            foreach (KeyValuePair<string,double> input in recipe.Input) // Specifying since it's not clear
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
