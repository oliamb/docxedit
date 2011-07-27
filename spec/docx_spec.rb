require 'docxedit'
require 'zip/zip'
require 'tmpdir'
require 'ruby-debug'

describe "DocxEdit::Docx", "#score" do
  
  before :each do
    @tmpdir = Dir.mktmpdir
    FileUtils.copy(File.join(File.dirname(__FILE__), 'fixtures/Archive4.docx'), @tmpdir)
    @doc = DocxEdit::Docx.new(File.join(@tmpdir, 'Archive4.docx'))
  end
  
  it "open the docx document" do
    @doc.zip_file.should_not be_nil
    @doc.zip_file.kind_of?(Zip::ZipFile).should be_true
  end
  
  it "find a given text string" do
    @doc.contains?(/\[WEEKLY_REPORT_WEEK_PARAGRAPH\]/).should be_true
    @doc.contains?(/inexistant string/).should be_false
  end
  
  it "replace a given string" do
    @doc.replace(/\[WEEKLY_REPORT_WEEK_PARAGRAPH\]/, "Another replacement string")
    @doc.contains?(/Another replacement string/).should be_true
    @doc.contains?(/\[WEEKLY_REPORT_WEEK_PARAGRAPH\]/).should be_false
  end
  
  it "replace a given string in header" do
    @doc.replace(/\[BENEFICIARY_FULL_NAME\]/, "Another replacement string")
    @doc.contains?(/Another replacement string/).should be_true
    @doc.contains?(/\[BENEFICIARY_FULL_NAME\]/).should be_false
  end
  
  it "commit change to file" do
    @doc.replace(/\[WEEKLY_REPORT_WEEK_PARAGRAPH\]/, "Another replacement string")
    @doc.contains?(/Another replacement string/).should be_true
    @doc.contains?(/\[WEEKLY_REPORT_WEEK_PARAGRAPH\]/).should be_false
    
    @doc.commit
    
    doc2 = DocxEdit::Docx.new(File.join(@tmpdir, 'Archive4.docx'))
    doc2.contains?(/Another replacement string/).should(be_true, doc2.xml_document.to_s)
    doc2.contains?(/\[WEEKLY_REPORT_WEEK_PARAGRAPH\]/).should be_false
  end

  it "commits changes to file at a new location, leaving the old version alone" do
    @doc.replace(/\[WEEKLY_REPORT_WEEK_PARAGRAPH\]/, "Another replacement string")
    @doc.contains?(/Another replacement string/).should be_true
    @doc.contains?(/\[WEEKLY_REPORT_WEEK_PARAGRAPH\]/).should be_false

    @doc.commit(File.join(@tmpdir, 'Archive4Copy.docx'))

    doc2 = DocxEdit::Docx.new(File.join(@tmpdir, 'Archive4Copy.docx'))
    doc2.contains?(/Another replacement string/).should(be_true, doc2.xml_document.to_s)
    doc2.contains?(/\[WEEKLY_REPORT_WEEK_PARAGRAPH\]/).should be_false

    doc1 = DocxEdit::Docx.new(File.join(@tmpdir, 'Archive4.docx'))
    doc1.contains?(/Another replacement string/).should(be_false, doc1.xml_document.to_s)
    doc1.contains?(/\[WEEKLY_REPORT_WEEK_PARAGRAPH\]/).should be_true

  end
  
  it "commit change to header" do
    @doc.replace(/\[BENEFICIARY_FULL_NAME\]/, "Another replacement string")
    @doc.contains?(/Another replacement string/).should be_true
    @doc.contains?(/\[BENEFICIARY_FULL_NAME\]/).should be_false
    
    @doc.commit
    
    doc2 = DocxEdit::Docx.new(File.join(@tmpdir, 'Archive4.docx'))
    doc2.contains?(/Another replacement string/).should(be_true, doc2.xml_headers[0].to_s)
    doc2.contains?(/\[BENEFICIARY_FULL_NAME\]/).should be_false
  end
  
  it "commit change whith a bigger end file size" do
    REXML::XPath.first(@doc.xml_document, "/*/*") << (REXML::Document.new "<root><p>A very long XML content</p><p>A very long XML content</p><p>A very long XML content</p><p>A very long XML content</p><p>A very long XML content</p><p>A very long XML content</p><p>A very long XML content</p><p>A very long XML content</p><p>A very long XML content</p><p>A very long XML content</p><p>A very long XML content</p><p>A very long XML content</p><p>A very long XML content</p><p>A very long XML content</p><p>A very long XML content</p><p>A very long XML content</p><p>A very long XML content</p><p>A very long XML content</p><p>A very long XML content</p></root>")
    @doc.commit
    doc2 = DocxEdit::Docx.new(File.join(@tmpdir, 'Archive4.docx'))
    doc2.contains?(/A very long XML content/).should(be_true, doc2.xml_document.to_s)
    doc2.contains?(/\[WEEKLY_REPORT_WEEK_PARAGRAPH\]/).should be_true
  end
  
  it "find a content block given its exact content string" do
    content = @doc.find_block_with_content("[WEEKLY_REPORT_WEEK_PARAGRAPH]")
    content.xml.should_not be_nil
    content.xml.to_s.should eql "<w:p w:rsidP='00297B5F' w:rsidR='00D12DFC' w:rsidRDefault='009B667E'><w:r><w:t>[WEEKLY_REPORT_WEEK_PARAGRAPH]</w:t></w:r><w:r w:rsidR='00D12DFC'><w:t xml:space='preserve'> </w:t></w:r></w:p>"
    content.content.should eql "[WEEKLY_REPORT_WEEK_PARAGRAPH]"
  end
  
  it "has an xml document attribute" do
    @doc.xml_document.should_not be_nil
  end
  
  it "Can insert a new block content before an existing node" do
    new_content = DocxEdit::ContentBlock.new(REXML::Document.new("<p>content</p>"), "content")
    content = @doc.find_block_with_content("[WEEKLY_REPORT_WEEK_PARAGRAPH]")
    @doc.insert_block(:before, content, new_content)
    REXML::XPath.first(@doc.xml_document, "//p[text()='content']").to_s.should eql "<p>content</p>"
  end
  
  it "Can insert multiple block consecutively" do
    new_content = DocxEdit::ContentBlock.new(REXML::Document.new("<p>content</p>"), "content")
    newer_content = DocxEdit::ContentBlock.new(REXML::Document.new("<p>newer content</p>"), "content")
    content = @doc.find_block_with_content("[WEEKLY_REPORT_WEEK_PARAGRAPH]")
    @doc.insert_block(:after, content, new_content)
    @doc.insert_block(:after, new_content, newer_content)
    REXML::XPath.first(@doc.xml_document, "//p[text()='content']").to_s.should eql "<p>content</p>"
    REXML::XPath.first(@doc.xml_document, "//p[text()='newer content']").to_s.should eql "<p>newer content</p>"
  end
  
  it "Can insert a new block content after an existing node" do
    new_content = DocxEdit::ContentBlock.new(REXML::Document.new("<p>content</p>"), "content")
    content = @doc.find_block_with_content("[WEEKLY_REPORT_WEEK_PARAGRAPH]")
    @doc.insert_block(:after, content, new_content)
    REXML::XPath.first(@doc.xml_document, "//p[text()='content']").to_s.should eql "<p>content</p>"
  end
  
  it "Can remove a block" do
    content = @doc.find_block_with_content("[WEEKLY_REPORT_WEEK_PARAGRAPH]")
    @doc.remove_block(content)
    REXML::XPath.first(@doc.xml_document, "//*[text()='[WEEKLY_REPORT_WEEK_PARAGRAPH]']").should be_nil
  end
end
