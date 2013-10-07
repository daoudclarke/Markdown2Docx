### ReleaseDocMarkdown 

This is an early stage gem that will allow you to create a template (currently used for release documents), and then apply some yaml that has variables defined containing markdown, or straight text.

Each variable corresponds to the text $/variable$.

To use:

require 'releaseDocMarkdown'
w = ReleaseDocMarkdown.open('templates/release.docx')
w.merge\_yaml('release.yaml')
w.save('release.docx')

To install, you might need to create the gem file:

gem build releaseDocMarkdown.gemspec
gem install .\releaseDocMarkdown-0.1.0.gem

Look in samples to see a sample yaml file.
