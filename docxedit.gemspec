# -*- encoding: utf-8 -*-

lib = File.expand_path('../lib/', __FILE__)
$:.unshift lib unless $:.include?(lib)
require 'docxedit/version'

spec = Gem::Specification.new do |s|
  s.name = 'docxedit'
  s.version = DocxEdit::VERSION
  s.platform    = Gem::Platform::RUBY
  
  s.summary = "Minimal Word Open Document format editor."
  s.description = %{minimal .docx file editor.}
  
  s.files = Dir['lib/**/*.rb'] + Dir['spec/**/*.spec'] + %w(LICENSE README)
  s.require_path = 'lib'
  
  s.add_dependency('rubyzip')
  s.add_development_dependency "rspec"
  
  
  s.author = "Olivier Amblet"
  s.email = "olivier@amblet.net"
  s.homepage = "http://github.com/oliamb/docxedit"
  
  s.required_rubygems_version = ">= 1.3.6"
  s.rubyforge_project         = "docxedit"
end