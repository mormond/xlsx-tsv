#
# Based on solution by Malyk @ StackOverflow
# https://stackoverflow.com/questions/54021712/how-to-properly-convert-an-xlsx-file-to-a-tsv-file-in-python
#

import pandas as pd
import os
import sys

if (len(sys.argv) > 1):
	# Assuming invocation "python3 convert_xlsx_to_tsv.py filename.xlsx"
	filename = sys.argv[1]

	# Split filename so we can generate .tsv version
	basefilename, ext = os.path.splitext(filename)

	# Load file into a Pandas dataframe
	exceldf = pd.read_excel(filename, 'Sheet1', index_col=None)

	# Replace all columns having spaces with underscores
	exceldf.columns = [c.replace(' ', '_') for c in exceldf.columns]

	# Replace all fields having line breaks with space
	df = exceldf.replace('\n', ' ', regex=True)

	#Write dataframe into csv
	df.to_csv(basefilename + '.tsv', sep='\t', encoding='utf-8',  index=False, quotechar='#', line_terminator='\r\n')

else:
	print ("Please call with the .xlsx filename as a command line argument.")