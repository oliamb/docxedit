spec = Gem::Specification.new do |s|
  s.name = 'docxedit'
  s.version = '0.0.1'
  s.summary = "Minimal Word Open Document format editor."
  s.description = %{Minimal Word Open Document format editor.}
  s.files = Dir['lib/**/*.rb'] + Dir['spec/**/*.spec']
  s.require_path = 'lib'
  s.autorequire = 'docxedit'
  s.has_rdoc = true
  s.extra_rdoc_files = Dir['[A-Z]*']
  s.rdoc_options << '--title' <<  'DocX Edit -- Minimal Word Open Document format editor'
  s.author = "Olivier Amblet"
  s.email = "olivier@amblet.net"
  s.homepage = "http://livingweb.ch"
end