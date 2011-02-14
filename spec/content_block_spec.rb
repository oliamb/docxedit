require 'docxedit'

describe "DocxEdit::ContentBlock", "#score" do  
  before :each do
    @xml = REXML::Document.new("<p><r>My Content</r></p>")
    @content = "My Content"
    @cb = DocxEdit::ContentBlock.new(@xml, @content)
  end
  
  it "has a xml attribute" do
    @cb.xml.should_not be_nil
    @cb.xml.should eql @xml
  end
  
  it "has a content attribute" do
    @cb.content.should eql @content
  end
  
  it "can update the content" do
    @cb.content = "Another content"
    @cb.xml.to_s.should eql REXML::Document.new("<p><r>Another content</r></p>").to_s
  end
  
  it "can be cloned" do
    new_cb = @cb.clone
    new_cb.content = "Another content"
    @cb.content.should eql @content
    
    REXML::XPath.first(new_cb.xml, "//r").should_not be_nil
    REXML::XPath.first(new_cb.xml, "//r").remove
    REXML::XPath.first(new_cb.xml, "//r").should be_nil
    REXML::XPath.first(@cb.xml, "//r").should_not be_nil
  end
end