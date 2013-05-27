#supprimmer extraction de documents.xml.rels
#mette id de l'image dans le html
#utiliser nokogiri pour trouver les noeuds images
#lire le fichier documents.xml.rels et faire les remplacement

require 'rubygems'
require 'zip/zip' # rubyzip gem
require 'nokogiri'

class WordXmlManipulate

  def self.open(path,destination, &block)
    self.new(path, destination,&block)
  end

  def initialize(path, destination, &block)
    if block_given?
      @zip = Zip::ZipFile.open(path)
	  yield(self)
      @zip.close
    else
      @zip = Zip::ZipFile.open(path)
	  
    end
	@destination=destination
  end
  
  def unzip_media_file (destination)
    puts 'extract and save img file'
	#unzip media folder
	@zip.each { |f|
     f_path=File.join(destination, File.basename(f.name))
	 #remove file if file exist in the directory
	 if File.exist?(f_path) then
       FileUtils.rm_rf f_path
     end
	 #extract file contain in media directory
	 if f.name=~/\/media\/+[\w+\-.]{1,}/ then
	   FileUtils.mkdir_p(File.dirname(f_path))
       @zip.extract(f, f_path)
	 end
   }
  end
  
  def convert_docx2html
    puts 'convertDocx2html'
    #unzip rels file
    @zip.extract('word/_rels/document.xml.rels', 'document.xml.rels'){true}
	
	xml = @zip.read('word/document.xml')
    #initialise Nokogiri reader with document.xml
	doc   = Nokogiri::XML(xml)

	#open xsl file
	xslt  = Nokogiri::XSLT(File.read('DocxHtml.xsl'))
	
	#convert document.xml in html using xsl sheet
	@html=xslt.transform(doc)
  end
  
  def correct_img_url(html, destination)
    puts 'correct img src'
    xml=@zip.read('word/_rels/document.xml.rels')
	doc = Nokogiri::XML(xml)
	#for each image
	html.css('img').each{ |node|
		#grab image url
		imgUrl= doc.css("Relationship[Id='#{node['alt']}']")[0]["Target"]
		#remove media before image name
		imgUrl=imgUrl.sub('media/','')
		#add dir url before image name
		imgUrl=File.join(destination, imgUrl)
		#replace node['href'] by imgUrl
		node['src']=imgUrl
	}
  end
  
  def add_meta_data(html)
    puts 'add meta data'
    #open core.xml wich contain modified and created date
    xml=@zip.read('docProps/core.xml')
	doc = Nokogiri::XML(xml)
	puts 'date created '+doc.xpath('//dcterms:created').first.text
	puts 'date modified '+doc.xpath('//dcterms:modified').first.text
	
	#select head node
	node = html.xpath('//html/head').first
    
	
	#add created date
	node_created = Nokogiri::XML::Node.new "meta", html
    node_created['content'] = doc.xpath('//dcterms:created').first.text
	node_created['name']='created'
    node.add_next_sibling(node_created)
	
	#add modified date
	node_modified = Nokogiri::XML::Node.new "meta", html
    node_modified['content'] = doc.xpath('//dcterms:modified').first.text
	node_modified['name']='modified'
    node.add_next_sibling(node_modified)
   
	#open app.xml wich contains page number
	xml=@zip.read('docProps/app.xml')
	doc = Nokogiri::XML(xml)
	puts 'number of pages '+doc.css('Pages').first.text
	
	#add page number
	node_pages = Nokogiri::XML::Node.new "meta", html
    node_pages['content'] = doc.css('Pages').first.text
	node_pages['name']='pages'
    node.add_next_sibling(node_pages)
	
	#add document name
	node_title = Nokogiri::XML::Node.new "title", html
    node_title.content = @zip.name
	node.add_next_sibling(node_title)
  end
  
  def save(html,destination)
    puts 'save html file'
    name =File.basename(@zip.name, '.docx')
	path=File.join(destination, name.concat('.html'))
    
    f = File.open(path, "w")
    f.write(html)
    f.close      
  end	
  
  def docx2html
	puts @destination
	
	html=self.convert_docx2html
	self.add_meta_data(html)
	self.correct_img_url(html,@destination)
	self.save(html,@destination)
	self.unzip_media_file(@destination)
	html
  end
end

if __FILE__ == $0
#  file = ARGV[0]
#  destination = ARGV[1] || File.dirname(__FILE__)
#  w = WordXmlManipulate.open(file)
#  html=w.convert2html
#  w.add_meta_data(html)
#  w.correct_img_url(html,destination)
  #save html
#  w.save(html,destination)
  
  #extract media folder for image
#  w.unzip_media_file(file, destination)
  w = WordXmlManipulate.open(ARGV[0],ARGV[1])
  html=w.docx2html
  
  puts"complete"
end