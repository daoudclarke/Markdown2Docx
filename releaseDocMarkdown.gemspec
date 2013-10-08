Gem::Specification.new do |s|
  s.name        = 'releaseDocMarkdown'
  s.version     = '0.1.1'
  s.date        = '2013-10-07'
  s.summary     = 'ReleaseDocMarkdown'
  s.description = 'Creates a release doc from a template and markdown in a yaml file'
  s.authors     = ['Dave Arkell', 'Daoud Clarke']
  s.email       = 'daoud.clarke@gorkana.com'
  s.files       = ['lib/releaseDocMarkdown.rb','samples/release.yaml','samples/release.docx']
  s.add_runtime_dependency 'dimensions'
  s.add_runtime_dependency 'zip'
  s.add_runtime_dependency 'nokogiri'
  s.add_runtime_dependency 'dimensions'
  s.add_runtime_dependency 'kramdown'
end
