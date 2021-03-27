import re
import csv
import string
import json
import pandas as pd
from difflib import SequenceMatcher

def clean(name):
	#converts the name into lowercase so that case sensativity isn't an issue
	cleaned = str(name).lower()
	#removes preposititions at,by,in,en,from and the characters following
	cleaned = re.sub(r"(\bat\b|\bby\b|\bin\b|\ben\b|\bfrom\b|\bor\b)+[\s\S]+", "", cleaned)
	#removes the word the
	cleaned = re.sub("\bthe\b", "", cleaned)
	#this removes all of the below listed words
	cleaned = re.sub(r"outlet|store|factory|retailer|restaurant", "" ,cleaned)
	#this matches text in parantheses which needs to be removed
	cleaned = re.sub(r"\([^)]*\)|(\s*,\s*)", "", cleaned)
	#this matches strings after slashes which need to be removed
	cleaned = re.sub(r"(\/|\\)+[\s\S]+", "", cleaned)
	#removes spaces
	cleaned = re.sub(r"\s", "", cleaned)
	#removes all special characters
	cleaned = re.sub(r"[^a-zA-Z0-9\s]", "", cleaned)
	return cleaned

#checks similarity of two words given that first part of words match
def similar1(a, b):
	a = str(a)
	b = str(b)
	#shortens both strings to length of the shorter of the two strings
	if(len(a) < len(b)):
		b = b[0:len(a)]
	else:
		a = a[0:len(b)]
	#compares the two strings of the same length for similarity
	return SequenceMatcher(None, a, b).ratio()
#checks similarity of two words given that second part of words match	
def similar2(a, b):
	a = str(a)
	b = str(b)
	#shortens both strings to length of the shorter of the two strings
	if(len(a) < len(b)):
		b = b[-len(a):]
	else:
		a = a[-len(b):]
	#compares the two strings of the same length for similarity
	return SequenceMatcher(None, a, b).ratio()

#read in the already classified list of tenants as a DataFrame
hist_list = pd.read_csv(open('Tenant Categories - Combined.csv'))
#input_list holds the values of the new tenants to be classified
input_list = pd.read_csv(open('input.csv'))
#branch categories with yelp mapping
#branches = pd.read_csv(open('Branch Categories.csv'))

#create an empty array to hold dictionaries which will be converted into a DataFrame
hold_hist_list = []
for index, row in hist_list.iterrows():
	hold_hist_list.append({'Clean':clean(row['Tenants']), 'Tenants':row['Tenants'], 'Branch 1':row['Branch 1'] , 'Branch 2':row['Branch 2'], 'Branch 3':row['Branch 3']})
#Adding the cleaned tenant names to the hist_list datatable
hist_list = pd.DataFrame(hold_hist_list)


######################This section of code checks for same and similar####################
hold_matches_list = []
hold_no_matches_list = []
#this will loop through the new list of tenants and compare them to hist_list
for index, row in input_list.iterrows():
	try:
		hold_row = hist_list[hist_list['Clean'] == clean(row['Tenants'])].iloc[0]
		hold_matches_list.append({'Tenants':row['Tenants'], 'Branch 1':hold_row['Branch 1'], 'Branch 2':hold_row['Branch 2'], 'Branch 3':hold_row['Branch 3']})
	
	except:
		#hold_score is used to keep track of the tenants that are most similar
		hold_score = 0
		#hold_tenant keeps the name of the tenant that is most similar to the new tenant we are looking up
		hold_tenant = ""
		for index, row_hist_list in hist_list.iterrows():
			#checks the first half of a name
			hold_similar1 = similar1(row['Tenants'],row_hist_list['Tenants'])
			#checks the second half of the name
			hold_similar2 = similar2(row['Tenants'],row_hist_list['Tenants'])
			temp_score = max(hold_similar1,hold_similar2) 
			if(temp_score > hold_score):
				hold_score = temp_score
				hold_tenant = row_hist_list['Tenants']
				
		if(hold_score > 0.75):
			hold_row = hist_list[hist_list['Tenants'] == hold_tenant].iloc[0]
			hold_matches_list.append({'Tenants':row['Tenants'], 'Branch 1':hold_row['Branch 1'], 'Branch 2':hold_row['Branch 2'], 'Branch 3':hold_row['Branch 3']})
		else:
			hold_no_matches_list.append({'Tenants': row['Tenants'], 'Location': row['Location']})
		
		
####################This section of code tries to determine branch###################
#for row in hold_no_matches_list:
#	yelp = yelp_lookup(row['Tenants'], row['Location'])
#	hold_score = 0
#	hold_index = 0
#	for index, row_branches in branches.iterrows():
#		current_score = (2 * len(yelp[0].intersection(set({row_branches['Yelp']})))) / (len(yelp[0]) + len(set({row_branches['Yelp']})))
#		if(current_score > hold_score):
#			hold_score = current_score
#			hold_index = index
#	if(hold_score > 0):
#		hold_row = branches.iloc[hold_index]
#		hold_matches_list.append({'Tenants':row['Tenants'], 'Branch 1':hold_row['Branch 1'], 'Branch 2':hold_row['Branch 2'], 'Branch 3':hold_row['Branch 3']})
		
hold_matches_list = pd.DataFrame(hold_matches_list)
hold_matches_list.to_csv("output.csv", index = False)
