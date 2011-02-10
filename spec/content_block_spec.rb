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
end