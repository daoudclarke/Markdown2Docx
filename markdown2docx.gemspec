Gem::Specification.new do |s|
  s.name        = 'markdown2docx'
  s.version     = '0.1.3'
  s.date        = '2013-10-07'
  s.summary     = 'markdown2docx'
  s.description = 'Combines markdown in a yaml file with a docx template'
  s.authors     = ['Dave Arkell', 'Daoud Clarke']
  s.email       = 'daoud.clarke@gmail.com'
  s.files       = ['lib/markdown2docx.rb','bin/md2docx','samples/release.yaml','samples/release.docx']
  s.executables = ['md2docx']
  s.add_runtime_dependency 'dimensions'
  s.add_runtime_dependency 'rubyzip', '~> 1.0.0'
  s.add_runtime_dependency 'nokogiri'
  s.add_runtime_dependency 'dimensions'
  s.add_runtime_dependency 'kramdown'
end
