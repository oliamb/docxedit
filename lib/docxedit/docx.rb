require 'zip/zip'
require "rexml/document"
require 'tempfile'
require_relative 'content_block'

module DocxEdit
  class Docx    
    attr_reader :zip_file, :xml_document, :xml_headers, :xml_footers
    
    def initialize(path, temp_dir=nil)
      @zip_file = Zip::ZipFile.new(path)
      @temp_dir = temp_dir
      bind_contents
    end
    
    def files
      return @xml_headers + @xml_footers + [@xml_document]
    end
    
    def contains?(text)
      files.each do |f|
        REXML::XPath.each f, XPATH_ALL_TEXT_NODE do |n|
          return true if n.text =~ text
        end
      end
      return false
    end
    
    # Persist changes in the Zip file
    def commit(new_path=nil)
      write_content(new_path)
    end
    
    def replace(reg_to_match, replacement)
      files.each do |f|
        REXML::XPath.each f, XPATH_ALL_TEXT_NODE do |n|
          n.text = n.text.gsub(reg_to_match, replacement)
        end
      end
    end
    
    def find_block_with_content(exact_content_string)
      files.each do
        node = REXML::XPath.first(@xml_document, "//w:p[descendant-or-self::*[text()='#{exact_content_string}']]")
        return ContentBlock.new(node, exact_content_string) unless node.nil?
      end
    end
    
    # insert the xml of a content block :before or :after the anchor_block
    def insert_block(position, anchor_block, new_block)
      case position
      when :before
        anchor_block.xml.previous_sibling = new_block.xml
      when :after
        anchor_block.xml.next_sibling = new_block.xml
      else
        raise "position argument must be one of :before, :after"
      end
    end
    
    def remove_block(block)
      block.xml.remove
    end
    
  private
    DOCUMENT_FILE_PATH = 'word/document.xml'
    XPATH_ALL_TEXT_NODE = "//*[text()]"
  
    def bind_contents
      @xml_document = REXML::Document.new(@zip_file.read(DOCUMENT_FILE_PATH))
      @xml_headers = []
      @xml_footers = []
      REXML::Document.new(@zip_file.read(DOCUMENT_FILE_PATH))
      
      idx = 1
      src = read_or_nil("word/header#{idx}.xml")
      while(!src.nil?) do
        @xml_headers << REXML::Document.new(src)
        idx = idx + 1
        src = read_or_nil("word/header#{idx}.xml")
      end
      
      idx = 1
      src = read_or_nil("word/footer#{idx}.xml")
      while(!src.nil?) do
        @xml_headers << REXML::Document.new(src)
        idx = idx + 1
        src = read_or_nil("word/footer#{idx}.xml")
      end
    end
    
    
    def read_or_nil(name)
      return @zip_file.read(name) rescue return nil
    end
    
    def write_entry(zip_output_stream, entry_name, xml_doc)
      zip_output_stream.put_next_entry(entry_name)
      output = ""
      xml_doc.write(output, 0)
      zip_output_stream.print output
    end
    
    def write_content(new_path=nil)
      if @temp_dir.nil?
        temp_file = Tempfile.new('docxedit-')
      else
        temp_file = Tempfile.new('docxedit-', @temp_dir)
      end
      Zip::ZipOutputStream.open(temp_file.path) do |zos|
        @zip_file.entries.each do |e|
          unless e.name == DOCUMENT_FILE_PATH || e.name =~ /word\/(header|footer)[0-9]+\.xml/
            zos.put_next_entry(e.name)
            zos.print e.get_input_stream.read
          end
        end
        
        write_entry(zos, DOCUMENT_FILE_PATH, @xml_document)
        (0 .. @xml_headers.size - 1).each do |idx|
          write_entry zos, "word/header#{idx + 1}.xml", @xml_headers[idx]
        end
        (0 .. @xml_footers.size - 1).each do |idx|
          write_entry zos, "word/footer#{idx + 1}.xml", @xml_footers[idx]
        end
      end
      
      if new_path.nil?
        path = @zip_file.name
        FileUtils.rm(path)
      else
        path = new_path
      end
      FileUtils.mv(temp_file.path, path)
      @zip_file = Zip::ZipFile.new(path)
    end
  end
end
