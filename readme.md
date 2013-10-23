### Markdown2Docx

This is an early stage gem that will allow you to create a template (currently used for release documents), and then apply some yaml that has variables defined containing markdown, or straight text.

Each variable corresponds to the text $/variable$.

To use:

```
sudo gem install markdown2docx
md2docx template.docx data.yaml output.docx
```

To use it within ruby code:

```ruby
require 'markdown2docx'
w = Markdown2Docx.open('templates/release.docx')
w.merge_yaml('release.yaml')
w.save('release.docx')
```

To install, from source, you'll need to create the gem file:

```
gem build markdown2docx.gemspec
gem install markdown2docx
```

Look in samples to see a sample yaml file.
