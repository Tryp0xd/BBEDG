# BBEDG - Bingo Brewer Excel Document Generator

## Welcome!
This tool was mainly produced for the [**Bingo Brewers**](https://discord.gg/BingoBrewers) splashers. This tool allows you to generate an Excel Spreadsheet which contains information about almost every single potion that you can brew, as well as miscellaneous ones like Wisp's Ice Flavoured Water.

## Usage
In order to successfully use this tool, you are required to install two python modules.
```py
import requests
import xlsxwriter
```
Both will be provided in a `requirements.txt` file in the GitHub repository.
Then you can run `pip install -r requirements.txt` to install the modules.

Once you installed both of the modules, you can run it as it is, or via command line.

This is one way I recommend running the program, and the steps are as followed:

1. Create an `input.in` file
2. Edit the `input.in` file with the inputs you want to enter. 
	format is as followed:
	```bash
	LM/DM # LM - Light mode, DM - Dark mode
	1/2/3 # 1 - Cheap Coffee, 2 - Decent Coffee, 3 - Black Coffee
	n
	... # The 'n' is price of an item that you changed in the miscPot dictionary,
		# The amount of items there are, you would put that many inputs.
		# As an example, this assumed it had only one item.
		# However if there were more potions, you'd put many more 'n' numbers
		# as input.
3. Run `python main.py < input.in > output.txt`

You can discard the output or not include it at all, but output is somewhat necessary if the spreadsheet is not being generated properly or just in case you included some debug prints yourself.

## Contact me
If any issues persist with the program, you can contact me via discord `trypo`
Let me know what the issue is by providing me the traceback via DMs (feel free to remove any private information you may have in that traceback but replace the private parts with (PRIVATE))
