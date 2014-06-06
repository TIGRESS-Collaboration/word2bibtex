#!/usr/bin/python
""" Python program to try to convert Microsoft Sources.xml to bibtex

This program by inspection only (not by any planning or design) reads the Sources.xml
file in the working directory and puts what ought to be a properly formatted bibtex
file to the standard output.

I didn't look up the proper structure or anything, I just tried to write this based on
looking at my own Sources.xml file and a sample bibtex file.

Roughly speaking the Sources.xml has a single Sources root element and a bunch
of Source sub-elements in that.  Then each Source sub-element has all the information
pertaining to that Source.

Each Source is parsed into a dictionary consisting of bibtex
fields and their content.  The article type and the argument for \cite commands
is also put into each dictionary.

Many of the Source sub-elements map directly to a bibtex field: eg:
<Title>Title of artice</Title> can map directly to   title={Title of artice},
These elements are identified in a top-level global dictionary object,
directTagMap.  For those tags, the text goes straight into the mapped
bibtex field.

The text of the ArticleType tag gives you the article type --
Journal, Book, etc.  This text needs to be translated to a valid
value to go after the @ sign in the bibtex entry.  This is also
handled by a global dictionary object, articleTypeMap.

Tags in ignoreMap are utterly ignored as they seem to only matter to
Word.

The Author tag is a particular troubling one.  As far as I can tell, Author can have
subelements Author, Corporate and Editor -- so Author can be a sub-element of Author.
Anyway, the second-level Author and Editor each have a single NameList sub-elements;
the NameList sub-element has a bunch of Person sub-elements; and each Person sub-element
has First, Middle, Last subelements.  The code for parsing these is particularly kludgy.

Finally, any tags which aren't processed otherwise, are added (both tag & text) to the
bibtex notes field so they can be manually handled.

This was written in Python 2.7.  It is NOT Unicode safe.

"""
from lxml import etree

# global definitions of valid things

# namespace dictionary.  Not sure how universal this is -- may need to use something other than 'b'?
# This was the namespace map in my own sources.xml file.  It would probably be smarter to 
# extract this straight from the XML file.

nsd = {'b': 'http://schemas.openxmlformats.org/officeDocument/2006/bibliography'}

# This is just some short code that handles a common thing I needed to do.
# If multiple values of a key exist, i wanted to just append them to the
# bibtex field.  If the field doesn't exist yet, create it; otherwise just
# add to it. 

def addToKey(dict,key,value, separator=', '):
	if not key in dict:
		dict[key]=str(value)
	else:
		dict[key]=dict[key] + separator + str(value)
	return dict  # I hope this works

# This is a map for which Source.txt tags map directly onto bibtex fields
# Used as "addto" rather than "set" 
directTagMap = { 	'Booktitle':'booktile',
					'City':'address',
					'Corporate':'organization',
					'CountryRegion':'address',
					'Edition':'edition',
					'Institution':'institution',
					'InternetSiteTitle':'howpublished',
					'Issue':'number',
					'JournalName':'journal',
					'Month':'month',
					'Pages':'pages',
					'PeriodcialTitle':'journal',
					'Publisher':'publisher',
					'StateProvince':'address',
					'Tag':'cite',
					'ThesisType':'type',
					'Title':'title',
					'URL':'HowPublished',
					'Volume':'volume',
					'Year':'year' }

ignoreMap = { 'Guid','LCID','RefOrder'}

#If not in these lists and not an Author, just add it to note

articleTypeMap = {  'ArticleInAPeriodical'		: 'article',
					'Book'						: 'book',
					'BookSection'				: 'inbook',
					'ConferenceProceedings'		: 'inproceedings', 
					'DocumentFromInternetSite'	: 'misc',
					'InternetSite'				: 'misc',
					'JournalArticle'			: 'article',
					'Misc'						: 'misc',
					'Report'					: 'techreport'	}

sourceList = []

# Load Sources file
doc = etree.parse("Sources.xml") # this new object is an ElementTree

# get root element
wordRoot = doc.getroot() #  this is an Element -- the root element of doc to be precise

# Get rid of the namespace prefix from all tags.  I'm not sure if this is really kosher
# but it made life easier for me ...
for elt in wordRoot.iter():  # this goes through ALL tags for now
	tnt=unicode(elt.tag)
	if tnt[0]=='{':
		tnt=tnt.rsplit('}',1)[-1]
	elt.tag = tnt
	
if wordRoot.tag != 'Sources':
	#print 'Root object is not Sources, it is %s' % wordRoot.tag
	raise Exception('SourcesFormatError')

# Now we should have only Source elements
n = 0
for wordSource in wordRoot.getchildren():
	if wordSource.tag!='Source':
		#print 'Non-Source child object in Sources: %s' % source.tag
		continue # don't make a fuss
	n = n + 1
	source = {} # Use a dictionary for the new source 
	#print "--Beginning of source %d" % n
	
	for elt in wordSource.getchildren(): # Iterate over child tags in "Source"
		tag=elt.tag
		text=elt.text
		if tag in ignoreMap:
			pass
		elif tag in directTagMap: # If this is a simple mapping
			field = directTagMap[tag]
			#print "----%s maps to %s with value %s" % ( tag, field, text )
			source = addToKey(source,field,text)
		elif tag == 'SourceType':   # Slightly more complex -- needs valuemapping
			if text in articleTypeMap:
				ATYPE = articleTypeMap[text]
			else:
				ATYPE = 'misc' # default
			#print "----SourceType %s maps to type %s" % (text,ATYPE)
			source = addToKey(source,'ATYPE',ATYPE)
			#fi
		elif tag=='Author':  # There should be an author within an author
			for wordAuth in elt.getchildren():
				authtype = wordAuth.tag # Ought to be editor or author
				if authtype == 'Editor':
					authtype = 'editor'
					#print '------editors found'
				else:
					authtype = 'author'
					#print '------authors found'
				#fi
				for wordNameList in wordAuth.getchildren():
					if wordNameList.tag=='Corporate':
						#print "----------Found corporate author, map to institution"
						source=addToKey(source,'institution',wordNameList.text)
					elif wordNameList.tag == 'NameList':
						etal=False # Set to "true" if I find a dummy name
						for person in wordNameList.getchildren():	
							if person.tag!='Person':
								pass
								#print "----------Expected Person, found %s -- skip" % person.tag
							else:
								eltl = person.find('Last')
								#print "----------Person found, eltl =",eltl
								if eltl!=None: # If a last name is found
									if eltl.text in ['a','a;','aaa','b.','et al.', 'et al',
														'b','c','d','e','x' ]:
										#print '------------Dummy name found, skip'
										etal = True
										continue
									name = '{'+eltl.text+'}'
									eltm = person.find('Middle')
									if eltm!=None:
										name = eltm.text + ' ' + name	
									eltf = person.find('First')
									if eltf!=None:
										name = eltf.text + ' ' + name
									# Full name ought to be assembled at this point
									source=addToKey(source,authtype,name,separator=' and ')
								else:								
									pass # print '--------------No last name found, skip'
								#fi eltl:
							#fi person.tag!='Person'
						#rof person in whatnot
						if etal: # If a dummy name was found and skipped,
							source=addToKey(source,authtype,'{{et al.}}',separator=' and ')
					else:  
						#print "--------Expected NameList, found %s -- skip" % wordNameList.tag
						continue
					#fi wordNameList.tag
				#rof wordNameList
			#rof wordAuth		
		else:
			source = addToKey(source,'note',"%s : %s" % (tag,text) )
		#fi
	#rof
	# Check if no author found
	if not 'author' in source:
		source['author']='Anyonymous'
	#print "--End of source %d" % n
	#print source
	sourceList.append(source)
#rof

#print sourceList

# Now output this as a bibtex formatted file
# To standard output for now, I'm afraid.

#print '---------------Bibtex starts here-----------------'
tagno = 0 # For use in untagged items
for source in sourceList:
	ATYPE=source.pop('ATYPE')
	if ATYPE==0:
		del ATYPE
		ATYPE='misc'
	cite = source.pop('cite')
	if cite==0:
		del cite
		cite='Unkn%4d' % tagno
		tagno = tagno + 1
	print ('@%s{%s' % (ATYPE,cite)),
	for field, value in source.iteritems():
		print ','
		print ('  %s = {%s}' % (field,value)),
	print ','
	print '}'
	print 
#rof source in sourceList	
	

			



		