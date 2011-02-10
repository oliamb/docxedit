require 'zip/zip'
require "rexml/document"
require_relative 'content_block'

module DocxEdit
  class Docx    
    attr_reader :zip_file, :xml_document
    
    def initialize(path)
      @zip_file = Zip::ZipFile.new(path)
      @xml_document = REXML::Document.new(read_content)
    end
    
    def contains?(text)
      REXML::XPath.each @xml_document, XPATH_ALL_TEXT_NODE do |n|
        return true if n.text =~ text
      end
      return false
    end
    
    # Persist changes in the Zip file
    def commit
      write_content
    end
    
    def replace(reg_to_match, replacement)
      REXML::XPath.each @xml_document, XPATH_ALL_TEXT_NODE do |n|
        n.text = n.text.gsub(reg_to_match, replacement)
      end
    end
    
    def find_block_with_content(exact_content_string)
      node = REXML::XPath.first(@xml_document, "//w:p[descendant-or-self::*[text()='#{exact_content_string}']]")
      return ContentBlock.new(node, exact_content_string)
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
  
    def read_content()
      return text = @zip_file.read(DOCUMENT_FILE_PATH)
    end
    
    def write_content()
      @zip_file.get_output_stream(DOCUMENT_FILE_PATH) do |input|
        output = ""
        @xml_document.write(output, 0)
        input.write output
      end
      @zip_file.commit
    end
  end
end