require 'rubygems'
require 'zip/zip' # rubyzip gem
require 'nokogiri'

class WordXmlManipulate
  def self.open(path, &block)
    self.new(path, &block)
  end

  def initialize(path, &block)
    if block_given?
      @zip = Zip::ZipFile.open(path)
      yield(self)
      @zip.close
    else
      @zip = Zip::ZipFile.open(path)
    end
  end
  
  def unzip_file (file, destination)
	#unzip rels file
  @zip.extract('word/_rels/document.xml.rels', 'document.xml.rels'){true}
	
	#unzip media directorie
	@zip.each { |f|
     f_path=File.join(destination, f.name.sub(/word\//,''))
	 if File.exist?(f_path) then
       FileUtils.rm_rf f_path
     end
	 if f.name=~/\/media\/+[\w+\-.]{1,}/ then
	   puts f.name
	   puts f_path
       FileUtils.mkdir_p(File.dirname(f_path))
       @zip.extract(f, f_path)
	 end
   }
  end
  
  def extract
    xml = @zip.read('word/document.xml')
    #initialise Nokogiri reader with document.xml
	doc   = Nokogiri::XML(xml)

	#open xsl file
	xslt  = Nokogiri::XSLT(File.read('DocxHtml.xsl'))
	
	#convert document.xml in html using xsl sheet
	xslt.transform(doc)
  end
  
  def show(html)
    puts html
  end
  
  def save(html,path)
    f = File.open(path, "w")
    f.write(html)
    f.close      
  end	  
end

if __FILE__ == $0
  file = ARGV[0]
  destination = ARGV[1] || File.dirname(__FILE__)
  save_file = File.join(destination, file.sub(/\.docx/, '.html'))
  w = WordXmlManipulate.open(file)
  w.unzip_file(file, destination)
  html=w.extract
  #w.show html
  puts save_file
  w.save(html,save_file)
  puts"complete"
end