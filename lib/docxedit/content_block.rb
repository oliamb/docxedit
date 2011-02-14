module DocxEdit
  class ContentBlock
    attr_accessor :content
    
    def initialize(xml, content)
      @xml = xml
      @content_key = content
      @content = content
    end
    
    def clone
      result = ContentBlock.new(rclone_xml(@xml), @content_key)
    end
    
    def xml
      unless(@content_key == @content)
        @node = REXML::XPath.first(@xml, "//*[text()='#{@content_key}']]")
        @node.text = @content
        @content_key = @content
      end
      return @xml
    end
    
  private
  
    def rclone_xml(node)
      new_node = node.clone
      unless node.node_type == :text
        node.children.each do |c|
          new_node << rclone_xml(c)
        end
      end
      return new_node
    end
  end
end