require 'nokogiri'

doc   = Nokogiri::XML(File.read('document.xml'))

xslt  = Nokogiri::XSLT(File.read('DocxHtmlv2.xsl'))
tmp =xslt.transform(doc)
puts tmp

f = File.open("html.xml", "w")
f.write(tmp)
f.close