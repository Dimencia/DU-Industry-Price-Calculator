# DU-Industry-Price-Calculator
Generates a spreadsheet that calculates the prices of end-products for Industry in Dual Universe

You can input the current market price of ores, and determine whether or not producing a given item is more profitable than just selling the ores (You still have to check the market for the price of whatever you want to produce)

It includes a 'With Talents' page, which assumes you have all Production talents maxed for all steps of the process, and takes into account any input cost reductions and output increases - many elements are only profitable with full talents.

It does not account for batching; it is assumed that if you are producing an item in mass, batching will be effectively canceled out - it provides the price of a single unit of whatever you're making, and may include ore values of less than 1 due to batching

You can edit the ore prices to match current market values and everything will update accordingly.



If you'd like to help with this, we could use all the new recipes in JSON format, matching the format of recipes_object_version.json

We don't actually care about the Time field and I don't know how you get it.  Also, Price, PriceToMake, and PriceToMakeWithTalents are obsolete... I didn't mean to leave the last two in that json.  But Price is also only really applicable to Ores.  You can fill it in if you want but I don't think I'll ever use it, it would change too rapidly

And, just fair warning, I had to add Uncommon Power Transformer M and S to the recipes, and don't know how to get the craft time so.  The Time is inaccurate on those two, if you use the recipes json
