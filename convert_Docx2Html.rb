require 'nokogiri'

doc   = Nokogiri::XML(File.read('word/document.xml'))

xslt  = Nokogiri::XSLT(File.read('DocxHtml.xsl'))
tmp =xslt.transform(doc)
puts tmp

f = File.open("page.html", "w")
f.write(tmp)
f.close