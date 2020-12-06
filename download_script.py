#!/usr/bin/python3
import requests
import urllib.request
from html.parser import HTMLParser
import os 
from bs4 import BeautifulSoup
import re




if not os.path.exists("excels"):
	os.mkdir("excels")

for x in range(2011,2016):																	#For each year in the requested range
	finlink = []
	urls = []
	url = "https://www.statistics.gr/el/statistics/-/publication/STO04/"+str(x)+"-Q4"		#I get the url with string formating
	html = urllib.request.urlopen(url).read()												#reading all the html in the page
	soup = BeautifulSoup(html)																#giving it to BS4 which formats it
	
	for link in soup.find_all('a'):															#search for all <a> tags in html
		urls.append(link.get("href"))														#get the href value(link) of the tag and append it to a list
		pass
	for link in urls:
		if re.search("VBZOni0vs5VJ",link):													# for each link in the list if it matches the pattern "VBZOni0vs5VJ" it's an excel file
			finlink.append(link)															# so append it to another list that only contains the excels
			pass
		pass
	if x==2015:																				#for 2015 the second was the third excel in the list so quick fix
		r = requests.get(finlink[1])
	else:
		r = requests.get(finlink[2])														#for the rest of the years it was always the second one
	d = r.headers['content-disposition']													#used to get the initial name of the excel file 
	fname = re.findall("filename=(.+)",d)[0]
	fname = fname.replace('"',"")															#fixing "" that appear because of greek language
	with open("excels/"+str(fname), 'wb') as fd:											#writing the file binarily chunk by chunk in the directory
		for chunk in r.iter_content(chunk_size=128):
			fd.write(chunk)
	fd.close()
			




