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
  Zip::ZipFile.open(file) { |zip_file|
   zip_file.each { |f|
     f_path=File.join(destination, f.name)
     FileUtils.mkdir_p(File.dirname(f_path))
     zip_file.extract(f, f_path) unless File.exist?(f_path)
   }
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
  save_file = ARGV[1] || file.sub(/\.docx/, '.html')
  w = WordXmlManipulate.open(file)
  html=w.extract
  w.extract
  w.show html
  w.save(html,save_file)
  puts"complete"
end