module DocxEdit
  class ContentBlock
    
    attr_accessor :content
    
    def initialize(xml, content)
      @xml = xml
      @content_key = content
      @content = content
    end
    
    def xml
      unless(@content_key == @content)
        @node = REXML::XPath.first(@xml, "//*[text()='#{@content_key}']]")
        @node.text = @content
        @content_key = @content
      end
      return @xml
    end
  end
end