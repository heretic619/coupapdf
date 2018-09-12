#!/usr/bin/env python3

import datetime
import argparse
from subprocess import Popen,PIPE
import urllib.parse
from xml.etree import ElementTree as ET
import csv
import os
import time
import pycurl
from io import BytesIO
from io import StringIO
from PyPDF2 import PdfFileReader, PdfFileMerger, PdfFileWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.pagesizes import landscape, portrait
from reportlab.platypus import Image
import math
from reportlab.lib.pagesizes import A4
from reportlab.lib.enums import TA_JUSTIFY
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib.enums import TA_RIGHT
from reportlab.pdfgen.canvas import Canvas
from reportlab.pdfbase.pdfmetrics import stringWidth
from shutil import copyfile
import smtplib
import traceback

#python3 pdfexporter.py -process csvgenerate -apikey {APIKEY} -site https://scottg-responsive.mindtouch.us -path /Test_Form_for_Autocomplete_Test -csvpath test.csv
#python3 pdfexporter.py -process pdfextract -apikey {APIKEY} -site https://scottg-responsive.mindtouch.us -csvpath test.csv -outfolder testfolder -pages 2
#python3 pdfexporter.py -process csvgenerate -apikey {APIKEY} -site https://success.coupa.com -path Support/Releases/19/New_Features -csvpath coupa.csv
#python3 pdfexporter.py -process pdfextract -apikey {APIKEY} -site https://success.coupa.com -csvpath coupa.csv -outfolder coupatestfolder -pages 1 -stylesheet guide
#python3 pdfexporter.py -process coupawholeshebang -apikey {APIKEY} -site https://success.coupa.com -path Support/Releases/19/New_Features -csvpath coupa.csv -stylesheet guide -outfolder coupatestfolder
#python3 pdfexporter.py -process coupawholeshebang -apikey {APIKEY} -site https://success.coupa.com -path Support/Releases/20 -csvpath coupa.csv -stylesheet guide -outfolder coupatestfolder
#python3 pdfexporter.py -process coupawholeshebang -username 'Customer Test' -password CoupaDocs2013! -site https://success.coupa.com -path Support/Docs/Power_Apps/Risk_Aware -csvpath coupa.csv -stylesheet guide -outfolder coupatestfolder

#python3 pdfexporter.py -process coupawholeshebang -username 'Customer Test' -password CoupaDocs2013! -site https://success.coupa.com -path Support/Docs/Power_Apps/Risk_Aware -csvpath coupa.csv -stylesheet guide -outfolder coupatestfolder -email "scottg@mindtouch.com" -title_long "Coupa Risk Aware Admin and User Guide" -title_short "Risk Aware Admin and User Guide" -doc_status "Risk Aware Release v1.0" -copyright "Copyright © 2018 Coupa Software, Inc. All rights reserved." -legal_message "Coupa reserves the right to make changes to this document and to its services. All trademarks are the property of their respective owners." -valid_date "Valid until May 31, 2018" -pageid_exclude 12455

def main():
	global backcover
	global title_long
	global title_short
	global doc_status
	global copyright
	global legal_message
	global valid_date
	global pageid_exclude
	parser = argparse.ArgumentParser(description='Process some integers.')
	parser.add_argument('-process', type=str, nargs='?',help='specify which process (csvgenerate or pdfextract or pdfcombine or coupacovermaker)')
	parser.add_argument('-username', type=str, nargs='?',help='username to autnethicate with')
	parser.add_argument('-password', type=str, nargs='?',help='password to autnethicate with')
	parser.add_argument('-apikey', type=str, nargs='?',help='apikey to autnethicate with')
	parser.add_argument('-site', type=str, nargs='?',help='site to work with')
	parser.add_argument('-path', type=str, nargs='?',help='path to generate csv file of page ids underneath or folder to look at')
	parser.add_argument('-pages', type=int, nargs='?',help='Number of pages per PDF file')
	parser.add_argument('-csvpath', type=str, nargs='?',help='Path and filename of the CSV file')
	parser.add_argument('-outfolder', type=str, nargs='?',help='Place to put the resulting files')
	parser.add_argument('-stylesheet', type=str, nargs='?',help='Stylesheet to use if using custom css stylesheet')
	parser.add_argument('-backcover', type=str, nargs='?',help='path of PDF file to append to back end of PDF')
	parser.add_argument('-email', type=str, nargs='?',help='email address to send success/failure to')
	parser.add_argument('-title_long', type=str,nargs='?',help='Coupa long title field')
	parser.add_argument('-title_short', type=str,nargs='?',help='Coupa short title field')
	parser.add_argument('-doc_status', type=str,nargs='?',help='Coupa doc status field')
	parser.add_argument('-copyright', type=str,nargs='?',help='Coupa copyright field')
	parser.add_argument('-legal_message', type=str,nargs='?',help='Coupa legal message field')
	parser.add_argument('-valid_date', type=str,nargs='?',help='Coupa valid date field')
	parser.add_argument('-pageid_exclude', type=str,nargs='?',help='Coupa excluded page id''s in CSV format')
	args = parser.parse_args()
	if not 'email' in args or args.email is None:
		toemail='scottg@mindtouch.com'
	else:
		toemail=args.email
	try:
		if not 'title_long' in args or args.title_long is None:
			title_long = "Coupa Risk Aware Admin and User Guide"
		else:
			title_long = args.title_long
		if not 'title_long' in args or args.title_long is None:
			title_short = "Risk Aware Admin and User Guide"
		else:
			title_short = args.title_short
		if not 'doc_status' in args or args.doc_status is None:
			doc_status = "Risk Aware Release v1.0"
		else:
			doc_status = args.doc_status
		if not 'copyright' in args or args.copyright is None:
			copyright = "Copyright © 2018 Coupa Software, Inc. All rights reserved."
		else:
			copyright = args.copyright
		if not 'legal_message' in args or args.legal_message is None:
			legal_message = "Coupa reserves the right to make changes to this document and to its services. All trademarks are the property of their respective owners."
		else:
			legal_message = args.legal_message
		if not 'valid_date' in args or args.valid_date is None:
			valid_date = "Valid until May 31, 2018"
		else:
			valid_date=args.valid_date
		if not 'pageid_exclude' in args or args.pageid_exclude is None:
			pageid_exclude = []
		else:
			tempp = args.pageid_exclude.split(',')
			pageid_exclude=[]
			for p in tempp:
				pageid_exclude=pageid_exclude+[int(p)]
		# print('excl')
		# print(pageid_exclude)


		# print(args)
		if not 'stylesheet' in args or args.stylesheet is None:
			args.stylesheet='default'
		if args.csvpath is None:
			csvpath='pdfexporter.csv'
		else:
			csvpath=args.csvpath

		if not 'backcover' in args or args.backcover is None:
			backcover=None
		else:
			backcover=args.backcover
		if 'apikey' in args and args.apikey is not None:
			# print('a')
			auth=args.apikey
		elif 'username' in args and args.username is not None:
			# print('b')
			auth=args.username+':'+args.password
		else:
			# print('c')
			auth=''
		# print('auth')
		# print(auth)
		# print(args.username)
		# print(args.password)
		if args.process=='csvgenerate':
			csvgenerate(args.site,args.path,args.csvpath,auth);
		elif args.process=='pdfextract':
			pdfextract(args.site,args.pages,csvpath,args.outfolder,auth,args.stylesheet,0);
		elif args.process=='apitester':
			apitester(args.site,auth)
		elif args.process=='tagstrip':
			tagstrip(args.site,csvpath,auth)
		elif args.process=='pdfcombine':
			pdfcombine(args.path)
		elif args.process=='coupacovermaker':
			coupacovermaker(args.path)
		elif args.process=='coupawholeshebang':
			outfolder=coupawholeshebang(args.site,args.path,csvpath,auth,args.stylesheet)
			sendemail("\nCoupa PDF Generated Successfully!  Location: "+outfolder,toemail)
			print("COMPLETE Location: "+outfolder)
	except Exception as e:
		tb = traceback.format_exc()
		sendemail("\nCoupa email failed!  Error: \n\n"+str(e)+"\n\n"+tb,toemail)
		print(str(e)+"\n\n"+tb)

def sendemail(msg,toemail):
	server = smtplib.SMTP('smtp.gmail.com', 587)
	server.starttls()
	server.login("sendsendingsent@gmail.com", 'WFu[Q6e"p/$x+ssk')
	server.sendmail("sendsendingsent@gmail.com", toemail, msg)
	server.quit()

def coupawholeshebang(site,path,csvpath,auth,stylesheet):
	tailoredoutfolder='coupapdf'+datetime.datetime.now().strftime("%Y%m%d%H%M%S")
	if not os.path.exists(tailoredoutfolder):
		os.makedirs(tailoredoutfolder)
	pages=1
	csvgenerate(site,path,csvpath,auth)
	coupacovermaker(os.path.abspath(tailoredoutfolder))
	outloc,bookmarks=coupapdfextract(site,pages,csvpath,tailoredoutfolder,auth,stylesheet,2)
	# print(bookmarks)
	coupaaddbookmarks(outloc,bookmarks)
	return(tailoredoutfolder)

def coupaaddbookmarks(inloc,bookmarks):
	output = PdfFileWriter()
	with open(inloc, 'rb') as inf:
		input1 = PdfFileReader(inf)
		#output.addPage(input1.getPage(0))
		output.appendPagesFromReader(input1)
		for p in bookmarks:
			#print(p)
			parent = output.addBookmark(p[1], p[0])
			# print('adding bookmark: '+str(p[1])+' on page '+str(p[0]))
			for c in p[2]:
				output.addBookmark(c[1], c[0], parent) # add child bookmark
				# print('adding bookmark: '+str(c[1])+' on page '+str(c[0]))
		outloc=inloc.replace('.pdf', 'bookmarked.pdf')
		with open(outloc,'wb') as outputfile:
			output.write(outputfile)

def coupacovermaker(folder):
	#Input variables
	global title_long
	global doc_status
	global copyright
	global legal_message
	#Make front cover
	# print(folder+"/aacoverpage.pdf")
	pdfmetrics.registerFont(TTFont('ArialBI', 'ARIALBI0.TTF'))
	pdfmetrics.registerFont(TTFont('ArialB', 'ARIALBD.TTF'))
	canvas = Canvas(folder+"/aacoverpage.pdf", pagesize=letter)
	page_width, page_height = canvas._pagesize
	wid=page_width
	canvas.drawImage('doc_cover_bg.png', 0, 0, width=wid,height=1.294117647058823*wid)

	canvas.setFont('ArialB', 42)
	canvas.setFillColorRGB(1,1,1)
	textWidth = stringWidth(title_long, 'ArialB', 42)
	x=0.5*page_width/8.5
	ytop=1.657*page_height/11
	ybot=1*page_height/11
	if 2*x+textWidth>=page_width:
		titlewords=title_long.split(' ')
		halfway=int(len(titlewords)/2)
		title1=' '.join(titlewords[0:halfway])
		title2=' '.join(titlewords[halfway:len(titlewords)])
		canvas.drawString(x, ytop, title1)
		canvas.drawString(x, ybot, title2)
	else:
		canvas.drawString(x, ybot, title_long)
	# print(textWidth)
	canvas.showPage()
	canvas.save()

	#Make legal page
	doc = SimpleDocTemplate(folder+"/aalegalpage.pdf",pagesize=letter,rightMargin=x,leftMargin=x,topMargin=72,bottomMargin=18)
	Story=[]
	styles=getSampleStyleSheet()
	styles.add(ParagraphStyle(name='Justify', alignment=TA_JUSTIFY))
	ptext = '<b>Document Status: </b>'+doc_status+'. '+valid_date+'.'
	Story.append(Spacer(1, 564))
	Story.append(Paragraph(ptext, styles["Justify"]))
	Story.append(Spacer(1, 12))
	Story.append(Paragraph(copyright, styles["Justify"]))
	Story.append(Spacer(1, 12))
	Story.append(Paragraph(legal_message, styles["Justify"]))
	Story.append(Spacer(1, 12))
	doc.build(Story)
	#Make back cover
	canvas = Canvas(folder+"/zzbackcoverpage.pdf", pagesize=letter)
	page_width, page_height = canvas._pagesize
	wid=page_width
	canvas.drawImage('doc_back_bg.png', 0, 0, width=wid,height=1.294117647058823*wid)

	canvas.showPage()
	canvas.save()

def pdfcombine(folder):
	global backcover
	if backcover is not None:
		copyfile(backcover, folder+'/zzzbackcover.pdf')
	# print('Combining PDFs...')
	merger = PdfFileMerger()
	subs=0
	subnum=0
	lotsapages=False
	for f in sorted(os.listdir(folder)):
		if '.pdf' in f and 'CombinedPages' not in f:
			# print(f)
			#zfile=open(folder+f,"rb")
			merger.append(folder+f)
			subs=subs+1
			#zfile.close()
		if subs==50:
			subs=0
			with open(folder+'/CombinedPages'+("%03d"%(subnum,))+'.pdf', 'wb') as fout:
				merger.write(fout)
			merger = PdfFileMerger()
			subnum=subnum+1
			lotsapages=True
	outloc=folder+'/CombinedPages'+("%03d"%(subnum,))+'.pdf'
	with open(outloc, 'wb') as fout:
		merger.write(fout)
	merger = PdfFileMerger()
	if lotsapages:
		for f in sorted(os.listdir(folder)):
			if '.pdf' in f and 'CombinedPages' in f:
				merger.append(folder+f)
		finaloutloc=folder+'CombinedPagesFinal.pdf'
		# print("COMPLETE " + folder + 'CombinedPagesFinal.pdf')
		with open(finaloutloc,'wb') as fout:
			merger.write(fout)
	else:
		finaloutloc=outloc
	return(finaloutloc)

def csvgenerate(site,startingpath,csvpath,auth):
	# print('csvgenerate')
	#o=curl(['c:\curl\curl.exe','-u',username+':'+password,site+'/@api/deki/pages/='+urllib.parse.quote_plus(urllib.parse.quote_plus(startingpath))])
	output=cget(site+'/@api/deki/pages/='+urllib.parse.quote_plus(urllib.parse.quote_plus(startingpath))+'?a=a',auth)
	# print(output)
	tree=ET.fromstring(output)
	topid=tree.attrib['id']
	csvlist,toppageslist=getRecursiveSubpages(topid,auth,site,'')
	with open(csvpath, 'w') as myfile:
		wr = csv.writer(myfile, quoting=csv.QUOTE_ALL)
		wr.writerow(csvlist)
		wr.writerow(toppageslist)
	# print(csvlist)

def pdfextract(site,pages,csvpath,outfolder,auth,stylesheet,tocoffset):
	outfolder=outfolder+datetime.datetime.now().strftime("%Y%m%d%H%M%S")
	with open(csvpath, 'r') as f:
		reader = csv.reader(f)
		zi=0
		zreader=[]
		for readrow in reader:
			# print(readrow)
			zreader.append(readrow)
			zi=zi+1
		pageids = zreader[0]
	if not os.path.exists(outfolder):
		os.makedirs(outfolder)
	if pages==1:
		ix=1
		for i in pageids:
			cfile(site+'/@api/deki/pages/'+i+'/pdf'+'?stylesheet='+stylesheet,auth,outfolder+'/page'+"%05d" % ix+'.pdf')
			ix=ix+1
	return(pdfcombine(outfolder+'/'))

def coupapdfextract(site,pages,csvpath,outfolder,auth,stylesheet,tocoffset):
	# print('coupapdfextract')
	toc=[]
	titles=[]
	pdflocs=[]
	chaplocs={}
	titlelimit=100
	with open(csvpath, 'r') as f:
		reader = csv.reader(f)
		# print('reader')
		# print(reader)
		zi=0
		zreader=[]
		for readrow in reader:
			# print(readrow)
			zreader.append(readrow)
			zi=zi+1
		pageids = zreader[0]
		# print('Page IDs')
		# print(pageids)
		toppages=zreader[1]
		# print('Top pages:')
		# print(toppages)
	if not os.path.exists(outfolder):
		os.makedirs(outfolder)
	if pages==1:
		numtblrows=len(pageids)+len(set(toppages))
		if numtblrows<=32:
			toclguess=1+tocoffset
		else:
			toclguess=math.ceil((numtblrows-32)/38)+1+tocoffset
		ix=1
		lastsec=''
		creattitlepage=False
		spills=[]
		for i,topp in zip(pageids,toppages):
			spillpages=0
			if topp!=lastsec:
				#create previous title page, except on first go
				if creattitlepage:
					chaptitle=outfolder+'/page'+"%05d" % nexttitlepg+'.pdf'
					canvas = Canvas(chaptitle, pagesize=letter)
					chaplocs[topp]=chaptitle
					page_width, page_height = canvas._pagesize
					wid=page_width
					canvas.drawImage('doc_chapter_bg.png', 0, 0, width=wid,height=1.294117647058823*wid)
					pdfmetrics.registerFont(TTFont('ArialBI', 'ARIALBI0.TTF'))
					pdfmetrics.registerFont(TTFont('ArialB', 'ARIALBD.TTF'))
					pdfmetrics.registerFont(TTFont('Arial', 'Arial.ttf'))
					canvas.setFont('ArialB', 28)
					canvas.setFillColorRGB(1,1,1)
					textWidth = stringWidth(prevtopp, 'ArialB', 28)
					x=0.5*page_width/8.5
					ytop=page_height-0.96*page_height/11
					ybot=page_height-1.5*page_height/11
					if 2*x+textWidth>=page_width:
						titlewords=prevtopp.split(' ')
						halfway=int(len(titlewords)/2)
						title1=' '.join(titlewords[0:halfway])
						title2=' '.join(titlewords[halfway+1:len(titlewords)])
						canvas.drawString(x, ytop, title1)
						canvas.drawString(x, ybot, title2)
					else:
						# print('topp')
						# print(prevtopp)
						# print(x)
						# print(ybot)
						canvas.drawString(x, ybot, prevtopp)
					y=page_height-2.2*page_height/11
					canvas.setFont('Arial', 10)
					canvas.setFillColorRGB(0,0,0)
					canvas.drawString(x, y, 'This chapter contains the following topics:')
					y=y-0.4*page_height/11
					
					for t,s in zip(subtitles,subsums):
						if y<50:
							spillpages=spillpages+1
							#new page
							canvas.showPage()
							canvas.save()
							canvas = Canvas(outfolder+'/page'+"%05d" % nexttitlepg+'n'+str(spillpages)+'.pdf', pagesize=letter)
							page_width, page_height = canvas._pagesize
							y=page_height-1.2*page_height/11
						canvas.setFont('ArialB', 12)
						canvas.setFillColorRGB(0.2,0.568,0.792)
						canvas.drawString(x, y, t)
						y=y-0.22*page_height/11
						canvas.setFont('Arial', 10)
						canvas.setFillColorRGB(0,0,0)

						remainingtext=s
						textWidth = stringWidth(remainingtext, 'Arial', 10)
						#ytop=1.54*page_height/11
						#ybot=1*page_height/11
						while textWidth>=page_width*7.5/8.5:
							#figure out how far over it is
							percentoff=textWidth/(page_width*7.5/8.5)
							#adjust it for error
							while True:
								percentoff=0.9*percentoff
								# print(percentoff)
								#pull the first
								textpull=remainingtext[0:int(len(remainingtext)*percentoff)].rsplit(' ', 1)[0]
								if stringWidth(textpull, 'Arial', 10)<page_width*7.5/8.5:
									break
							canvas.drawString(x, y, textpull)
							y=y-0.2*page_height/11
							remainingtext=remainingtext.replace(textpull, '', 1).lstrip(' ')
							textWidth = stringWidth(remainingtext, 'Arial', 10)
						canvas.drawString(x, y, remainingtext)
						y=y-0.4*page_height/11
					spills[lastspillindex]=spillpages
					# print(textWidth)
					canvas.showPage()
					canvas.save()
				#set aside page number for next title page
				lastspillindex=len(spills)
				nexttitlepg=ix
				prevtopp=topp
				ix=ix+1
				ix=ix+1
				lastsec=topp
				creattitlepage=True
				subtitles=[]
				subsums=[]
			spills=spills+[0]

			# print(i)
			#o=curl(['c:\curl\curl.exe','-u',username+':'+password,'--output',outfolder+'/page'+i+'.pdf',site+'/@api/deki/pages/'+i+'/pdf'])
			pdfloc=outfolder+'/page'+"%05d" % ix+'.pdf'
			pdflocs=pdflocs+[pdfloc]
			cfile(site+'/@api/deki/pages/'+i+'/pdf'+'?stylesheet='+stylesheet,auth,pdfloc)
			#os.rename('page'+i+'.pdf','c:/users/jscot/Desktop/testfolder/page'+i+'.pdf')
			with open(pdfloc,'rb') as pdffile:
				pdf = PdfFileReader(pdffile)
				pdfn=pdf.getNumPages()
				# print('PDF pages:')
				# print(pdfn)
				output=cget(site+'/@api/deki/pages/'+i+'?a=a',auth)
				tree=ET.fromstring(output)
				title=charclean(tree.find('title').text)
				try:
					summary=charclean(tree.find('.//properties/property[@name="mindtouch.page#overview"]/contents').text)
				except:
					summary=''
				subtitles.append(title)
				subsums.append(summary)
				# print(subsums)
				toc=toc+[pdfn]
				titles=titles+[title]
				if len(title)>titlelimit:
					title=title[:titlelimit]+'..'
				#ptext = title+'. . . . . . . . .'+str(pagesum)
				#Story.append(Paragraph(ptext, styles["Justify"]))
				#Story.append(Spacer(1, 12))
			ix=ix+1


			#create last chapter page
			if creattitlepage:

				chaptitle=outfolder+'/page'+"%05d" % nexttitlepg+'.pdf'
				chaplocs[topp]=chaptitle
				canvas = Canvas(chaptitle, pagesize=letter)
				page_width, page_height = canvas._pagesize
				wid=page_width
				canvas.drawImage('doc_chapter_bg.png', 0, 0, width=wid,height=1.294117647058823*wid)
				pdfmetrics.registerFont(TTFont('ArialBI', 'ARIALBI0.TTF'))
				pdfmetrics.registerFont(TTFont('ArialB', 'ARIALBD.TTF'))
				pdfmetrics.registerFont(TTFont('Arial', 'Arial.ttf'))
				canvas.setFont('ArialB', 28)
				canvas.setFillColorRGB(1,1,1)
				textWidth = stringWidth(prevtopp, 'ArialB', 28)
				x=0.5*page_width/8.5
				ytop=page_height-0.96*page_height/11
				ybot=page_height-1.5*page_height/11
				if 2*x+textWidth>=page_width:
					titlewords=prevtopp.split(' ')
					halfway=int(len(titlewords)/2)
					title1=' '.join(titlewords[0:halfway])
					title2=' '.join(titlewords[halfway+1:len(titlewords)])
					canvas.drawString(x, ytop, title1)
					canvas.drawString(x, ybot, title2)
				else:
					# print('topp')
					# print(prevtopp)
					# print(x)
					# print(ybot)
					canvas.drawString(x, ybot, prevtopp)
				y=page_height-2.2*page_height/11
				canvas.setFont('Arial', 10)
				canvas.setFillColorRGB(0,0,0)
				canvas.drawString(x, y, 'This chapter contains the following topics:')
				y=y-0.4*page_height/11
				
				for t,s in zip(subtitles,subsums):
					if y<25:
						spillpages=spillpages+1
						#new page
						canvas.showPage()
						canvas.save()
						canvas = Canvas(outfolder+'/page'+"%05d" % nexttitlepg+'n'+str(spillpages)+'.pdf', pagesize=letter)
						page_width, page_height = canvas._pagesize
						y=page_height-1.2*page_height/11
					canvas.setFont('ArialB', 12)
					canvas.setFillColorRGB(0.2,0.568,0.792)
					canvas.drawString(x, y, t)
					y=y-0.22*page_height/11
					canvas.setFont('Arial', 10)
					canvas.setFillColorRGB(0,0,0)

					remainingtext=s
					textWidth = stringWidth(remainingtext, 'Arial', 10)
					#ytop=1.54*page_height/11
					#ybot=1*page_height/11
					while textWidth>=page_width*7.5/8.5:
						#figure out how far over it is
						percentoff=textWidth/(page_width*7.5/8.5)
						#adjust it for error
						while True:
							percentoff=0.9*percentoff
							# print(percentoff)
							#pull the first
							textpull=remainingtext[0:int(len(remainingtext)*percentoff)].rsplit(' ', 1)[0]
							if stringWidth(textpull, 'Arial', 10)<page_width*7.5/8.5:
								break
						canvas.drawString(x, y, textpull)
						y=y-0.2*page_height/11
						remainingtext=remainingtext.replace(textpull, '', 1).lstrip(' ')
						textWidth = stringWidth(remainingtext, 'Arial', 10)
					canvas.drawString(x, y, remainingtext)
					y=y-0.4*page_height/11
				spills[lastspillindex]=spillpages
				 
				canvas.showPage()
				canvas.save()


		bookmarks=makepdftoc(titlelimit,pageids,toppages,titles,toc,spills,pdflocs,chaplocs,outfolder,site,auth,stylesheet,toclguess,False)
		if os.path.isfile(outfolder+"/atoc2.pdf"):
			with open(outfolder+"/atoc2.pdf",'rb') as pdffile:
				pdf = PdfFileReader(pdffile)
				pdfn=pdf.getNumPages()
		else:
			pdfn=0
		bookmarks=makepdftoc(titlelimit,pageids,toppages,titles,toc,spills,pdflocs,chaplocs,outfolder,site,auth,stylesheet,pdfn+1+tocoffset,True)
	else:
		bookmarks=[]
		j=0
		stringme=''
		for i in pageids:
			stringme=stringme+','+i
			j=j+1
			if j==pages:
				# print(stringme)
				#output=curl(['c:\curl\curl.exe','-u',username+':'+password,'--output',outfolder+'/pdf'+str(time.time())+'.pdf',site+'/@api/deki/pages/book/?pageids='+stringme])
				cfile(site+'/@api/deki/pages/book/?pageids='+stringme+'&stylesheet='+stylesheet,auth,outfolder+'/pdf'+str(time.time())+'.pdf')
				j=0
				stringme=''
		# print(stringme)
		#output=curl(['c:\curl\curl.exe','-u',username+':'+password,'--output',outfolder+'/pdf'+str(time.time())+'.pdf',site+'/@api/deki/pages/book/?pageids='+stringme])
		cfile(site+'/@api/deki/pages/book/?pageids='+stringme,auth,outfolder+'/pdf'+str(time.time())+'.pdf')
	return(pdfcombine(outfolder+'/'),bookmarks)

def addCoupaTexttoPDF(inpdf,topleft,topright,bottomleft,startingpagen):
	output = PdfFileWriter()
	with open(inpdf, "rb") as inf:
		input1 = PdfFileReader(inf)
		npg=input1.getNumPages()
		# print("document1.pdf has %d pages." % npg)
		for p in range(0,npg):
			mpage=input1.getPage(p)
			input2=makeHF(topleft,topright,bottomleft,str(startingpagen+p))
			mpage.mergePage(input2.getPage(0))
			output.addPage(mpage)
		outpdf=inpdf.rstrip('.pdf')+'m.pdf'
		with open(outpdf, "wb") as outf:
			output.write(outf)
	os.remove(inpdf)

def makeHF(topleft,topright,bottomleft,bottomright):
	packet = BytesIO()
	canvas = Canvas(packet, pagesize=letter)
	page_width, page_height = canvas._pagesize
	canvas.setFont('ArialB', 10)
	canvas.setFillColorRGB(0.2,0.568,0.792)
	canvas.drawString(0.5*page_width/8.5, (11-0.5)*page_height/11, topleft)
	textWidth = stringWidth(topright, 'Arial', 10)
	canvas.drawString((8.5-0.5)*page_width/8.5-textWidth, (11-0.5)*page_height/11, topright)
	canvas.drawString(0.5*page_width/8.5, 0.5*page_height/11, bottomleft)

	textWidth = stringWidth(bottomright, 'Arial', 10)
	canvas.drawString((8.5-0.5)*page_width/8.5-textWidth, 0.5*page_height/11, bottomright)
	canvas.save()
	packet.seek(0)
	return(PdfFileReader(packet))

def makepdftoc(titlelimit,pageids,toppages,titles,toc,spills,pdflocs,chaplocs,outfolder,site,auth,stylesheet,tocl,headpages):
	global title_short
	global valid_date
	styles=getSampleStyleSheet()
	canvas = Canvas(outfolder+"/atoc.pdf", pagesize=letter)
	page_width, page_height = canvas._pagesize
	styles.add(ParagraphStyle(name='Justify', alignment=TA_JUSTIFY))
	styles.add(ParagraphStyle(name='TOCsection',parent=styles['Normal'],textColor='#3391CA',fontSize=12,fontName='ArialB'))
	styles.add(ParagraphStyle(name='TOCitem',parent=styles['Normal'],leftIndent=page_width/11/4,fontSize=10))
	styles.add(ParagraphStyle(name='right', parent=styles['Normal'], alignment=TA_RIGHT,fontSize=10))
	styles.add(ParagraphStyle(name='right_h', parent=styles['Normal'], alignment=TA_RIGHT,textColor='#3391CA',fontSize=12,fontName='ArialB'))
	canvas.drawImage('doc_chapter_bg.png', 0, 0, width=page_width,height=page_height)
	pdfmetrics.registerFont(TTFont('Arial', 'Arial.ttf'))
	canvas.setFont('ArialB', 28)
	canvas.setFillColorRGB(1,1,1)
	x=0.5*page_width/8.5
	rowsonfirstpage=34
	#x=72
	y=9.5*page_height/11
	canvas.drawString(x, y, 'Table of Contents')
	tbl_data = []
	ix=1
	buildadoc=False
	lastsec=''
	colw=[page_width-x*2-page_width/11,page_width/11]
	pagesum=1+tocl
	bookmarks=[]
	# print(spills)
	# print('spills: '+str(len(spills)))
	# print('toc: '+str(len(toc)))
	for title,pageid,topp,pdfn,spill,pdfloc in zip(titles,pageids,toppages,toc,spills,pdflocs):
		if topp!=lastsec:
			ix=ix+1
			if ix==rowsonfirstpage:
				tbl = Table(tbl_data,colw)
				w,h=tbl.wrapOn(canvas, page_width-2*x,page_height-2*x)
				tbl.drawOn(canvas, x,0.813*page_height-h)
				# print('###################################################')
				# print("Page height minus h: "+str(0.813*page_height-h))
				# print(page_height)
				# print(h)
				canvas.showPage()
				canvas.save()
				doc = SimpleDocTemplate(outfolder+"/atoc2.pdf",pagesize=letter,rightMargin=x,leftMargin=x,topMargin=72,bottomMargin=18)
				Story=[]
				tbl_data=[]
				buildadoc=True
			tbl_data.append([Paragraph(''+topp+'', styles["TOCsection"]), Paragraph(str(pagesum), styles['right_h'])])
			# print(chaplocs)
			try:
				#addCoupaTexttoPDF(chaplocs[topp],'','',valid_date,pagesum)
				addCoupaTexttoPDF(chaplocs[topp],'','','','')
			except:
				# print('File was modified')
				continue
			lastsec=topp
			bookmarks.append([pagesum-1,topp,[]])
			pagesum=pagesum+1+spill
		tbl_data.append([Paragraph(title, styles["TOCitem"]), Paragraph(str(pagesum), styles['right'])])
		bookmarks[-1][2].append([pagesum-1,title,[]])
		ix=ix+1
		if ix==rowsonfirstpage:
			tbl = Table(tbl_data,colw)
			w,h=tbl.wrapOn(canvas, page_width-2*x,page_height-2*x)
			tbl.drawOn(canvas, x,0.813*page_height-h)
			# print('###################################################')
			# print("Page height minus h: "+str(0.813*page_height-h))
			# print(page_height)
			# print(h)
			canvas.showPage()
			canvas.save()
			doc = SimpleDocTemplate(outfolder+"/atoc2.pdf",pagesize=letter,rightMargin=x,leftMargin=x,topMargin=72,bottomMargin=18)
			Story=[]
			tbl_data=[]
			buildadoc=True
		if (headpages):
			addCoupaTexttoPDF(pdfloc,title_short,topp,valid_date,pagesum)
		pagesum=pagesum+pdfn
	if buildadoc:
		tbl = Table(tbl_data,colw)
		Story.append(tbl)
		doc.build(Story)
	else:
		tbl = Table(tbl_data,colw)
		w,h=tbl.wrapOn(canvas, page_width-2*x,page_height-2*x)
		tbl.drawOn(canvas, x,0.813*page_height-h)
		# print('###################################################')
		# print("Page height minus h: "+str(0.813*page_height-h))
		# print(page_height)
		# print(h)
		canvas.showPage()
		canvas.save()
	return(bookmarks)

def coord(x, y, unit=1):
    x, y = x * unit, height -  y * unit
    return x, y

def apitester(site,auth):
	# print('apitester')
	#output=curl(['c:\curl\curl.exe','-u',username+':'+password,site+'/@api/deki/pages/'])
	output=cget(site+'/@api/deki/pages/'+'?a=a',auth)
	# print(output)
def tagstrip(site,csvpath,auth):
	# print('tagstrip')
	with open(csvpath, 'r') as f:
		reader = csv.reader(f)
		pageids = list(reader)[0]
	for i in pageids:
		# print(i)
		#output=curl(['c:\curl\curl.exe','-u',username+':'+password,site+'/@api/deki/pages/'+i+'/tags'])
		output=cget(site+'/@api/deki/pages'+i+'/tags?a=a',auth)
		tree=ET.fromstring(output)
		removearray=[]
		for child in tree:
			try:
				if 'article:' not in child.attrib['value']:
					#strip tag
					# print('stripping tag: '+child.attrib['value'])
					removearray.append(child)
			except:
				print("error would have happened...")
		for child in removearray:
			tree.remove(child)
		payload=ET.tostring(tree).decode('utf-8')
		#output=curl(['c:\curl\curl.exe','-X','PUT','-H','Content-Type: application/xml','-d',payload,'-u',username+':'+password,site+'/@api/deki/pages/'+i+'/tags'])
		outupt=cput(site+'/@api/deki/pages/'+i+'/tags?a=a',auth)

def charclean(strin):
	#if 'Â' in strin or 'â' in strin:
		# print(strin)
	return(strin.replace('Â\xa0',' ').replace('â\x80\x98',"'").replace('â\x80\x99',"'"))

def curl(arglist):
	# print(arglist)
	p = Popen(arglist, stdin=PIPE, stdout=PIPE, stderr=PIPE)
	output, err = p.communicate()
	return output
def cget(loc,auth):
	buffer = BytesIO()
	c = pycurl.Curl()
	c.setopt(c.URL, loc)
	if ':' in auth:
		#username and password
		#c.setopt(pycurl.HTTPAUTH, pycurl.HTTPAUTH_BASIC)
		c.setopt(pycurl.USERPWD,auth)
	else:
		#apikey
		loc=loc+'&apikey='+auth
	# print('cget: '+loc)
	c.setopt(c.URL, loc)
	c.setopt(c.WRITEDATA, buffer)
	#c.setopt(pycurl.USERPWD, '%s:%s' % (username, password))
	c.perform()
	c.close()
	body = buffer.getvalue()
	#print(body)
	return(body.decode('iso-8859-1'))
def cfile(loc,auth,path):
	buffer = BytesIO()
	c = pycurl.Curl()
	if ':' in auth:
		#username and password
		#c.setopt(pycurl.HTTPAUTH, pycurl.HTTPAUTH_BASIC)
		c.setopt(pycurl.USERPWD,auth)
	else:
		#apikey
		loc=loc+'&apikey='+auth
	c.setopt(c.URL, loc)
	# print(loc)
	# print(path)
	with open(path, 'wb') as f:
		c.setopt(c.WRITEFUNCTION, f.write)
		c.perform()
	c.close()
	return True
def getRecursiveSubpages(idin,auth,site,currenttop):
	#check if current page is a guide or category first
	o=cget(site+'/@api/deki/pages/'+idin+'?a=a',auth)
	tree=ET.fromstring(o)
	tags=tree.findall('.//tag')
	ztitle=tree.findall('.//title')[0].text
	inpdf=True
	for tag in tags:
		if tag.get('value')=='article:topic-category' or tag.get('value')=='article:topic-guide':
			inpdf=False
			currenttop=ztitle
	if inpdf and int(idin) not in pageid_exclude:
		pages=[int(idin)]
		toptitles=[currenttop]
	else:
		pages=[]
		toptitles=[]
	#o=curl(['c:\curl\curl.exe','-u',username+':'+password,site+'/@api/deki/pages/'+idin+'/subpages'])
	o=cget(site+'/@api/deki/pages/'+idin+'/subpages?a=a',auth)
	tree=ET.fromstring(o)
	#find each subpage
	thefs=tree.findall('page.subpage')
	for f in thefs:
		upages,utoptitles=getRecursiveSubpages(f.attrib['id'],auth,site,currenttop)
		pages = pages + upages
		toptitles=toptitles+utoptitles
	#join results into an array and return it
	return(pages,toptitles)

main()