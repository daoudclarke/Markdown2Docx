require 'rubygems'
require 'zip' # 'zip/zip' # rubyzip gem
require 'nokogiri'
require 'yaml'
require 'dimensions'
require 'kramdown'

YAML::ENGINE.yamler = 'syck'

class Markdown2Docx
  def self.open(path, &block)
    self.new(path, &block)
  end

  def initialize(path, &block)
    @replace = {}
    @media = {}
    if block_given?
      @zip = Zip::File.open(path)
      yield(self)
      @zip.close
    else
      @zip = Zip::File.open(path)
    end
  end

  def get_node(run, search)
    node_to_find = (run.xpath(".//#{search}")).first
    if node_to_find.nil?
      node_to_find = Nokogiri::XML::Node.new search, @doc
      run.add_child node_to_find
    end
    node_to_find
  end

  def get_emus_dimensions(image_file)
    dimensions = Dimensions.dimensions(image_file)
    px_width = dimensions[0]
    px_height = dimensions[1]
    dpi = 96.0
    emus_per_inch = 914400.0
    emus_per_cm = 360000.0
    max_width_cm = 15.0
    emus_width = px_width / dpi * emus_per_inch
    emus_height = px_height / dpi * emus_per_inch
    max_width = max_width_cm * emus_per_cm
    if (emus_width > max_width)
      ratio = emus_height / emus_width
      emus_width = max_width
      emus_height = emus_width * ratio
    end

    return emus_width.round, emus_height.round
  end

  def create_image_node(image_file, description)
    image_name = File.basename image_file
    docx_image_path = File.join 'media', image_name
    @media[File.join 'word',docx_image_path] = File.binread image_file
    rel_id = add_rels_link(docx_image_path, :image)
    check_for_jpeg_content_type
    drawing = Nokogiri::XML::Node.new 'w:drawing', @doc
    inline = Nokogiri::XML::Node.new 'wp:inline', @doc
    drawing.add_child inline
    inline['distT'] = '0'
    inline['distB'] = '0'
    inline['distL'] = '0'
    inline['distR'] = '0'

    extent = Nokogiri::XML::Node.new 'wp:extent', @doc
    inline.add_child extent
    emu_dimensions = get_emus_dimensions image_file
    extent['cx'] = emu_dimensions[0]
    extent['cy'] = emu_dimensions[1]

    effect_extent = Nokogiri::XML::Node.new 'wp:effectExtent', @doc
    inline.add_child effect_extent
    effect_extent['l'] = 0 # '19050'
    effect_extent['t'] = 0
    effect_extent['r'] = 0 #2099'
    effect_extent['b'] = 0

    docPr = Nokogiri::XML::Node.new 'wp:docPr', @doc
    inline.add_child docPr
    docPr['id'] = '1'
    docPr['name'] = image_name
    docPr['descr'] = description

    non_visual_graphic_props = Nokogiri::XML::Node.new 'wp:cNvGraphicFramePr', @doc
    inline.add_child non_visual_graphic_props
    frame_locks = Nokogiri::XML::Node.new 'a:graphicFrameLocks', @doc
    frame_locks.add_namespace_definition 'a', 'http://schemas.openxmlformats.org/drawingml/2006/main'
    frame_locks['noChangeAspect'] = '1'
    non_visual_graphic_props.add_child frame_locks

    graphic = Nokogiri::XML::Node.new 'a:graphic', @doc
    inline.add_child graphic
    graphic.add_namespace_definition 'a', 'http://schemas.openxmlformats.org/drawingml/2006/main'

    graphic_data = Nokogiri::XML::Node.new 'a:graphicData', @doc
    graphic.add_child graphic_data
    graphic_data['uri'] = 'http://schemas.openxmlformats.org/drawingml/2006/picture'

    pic = Nokogiri::XML::Node.new 'pic:pic', @doc
    pic.add_namespace_definition 'pic', 'http://schemas.openxmlformats.org/drawingml/2006/picture'
    graphic_data.add_child pic

    nv_picture_properties = Nokogiri::XML::Node.new 'pic:nvPicPr', @doc
    pic.add_child nv_picture_properties
    nv_drawing_properties =  Nokogiri::XML::Node.new 'pic:cNvPr', @doc
    nv_drawing_properties['id'] = '0'
    nv_drawing_properties['name'] = image_name
    nv_drawing_properties['descr'] = description
    nv_picture_properties.add_child nv_drawing_properties

    nv_picture_drawing_properties =  Nokogiri::XML::Node.new 'pic:cNvPicPr', @doc
    nv_picture_properties.add_child nv_picture_drawing_properties
    pic_locks =  Nokogiri::XML::Node.new 'a:picLocks', @doc
    pic_locks['noChangeAspect'] = '1'
    nv_picture_drawing_properties.add_child pic_locks

    blip_fill = Nokogiri::XML::Node.new 'pic:blipFill', @doc
    pic.add_child blip_fill

    blip = Nokogiri::XML::Node.new 'a:blip', @doc
    blip_fill.add_child blip
    blip['r:embed'] = "rId#{rel_id}"
    blip['cstate'] = 'print'

    src_rect =  Nokogiri::XML::Node.new 'a:srcRect', @doc
    blip_fill.add_child src_rect

    stretch =  Nokogiri::XML::Node.new 'a:stretch', @doc
    fill_rect =  Nokogiri::XML::Node.new 'a:fillRect', @doc
    stretch.add_child fill_rect
    blip_fill.add_child stretch

    shape_properties =  Nokogiri::XML::Node.new 'pic:spPr', @doc
    shape_properties['bwMode'] = 'auto'
    transform =  Nokogiri::XML::Node.new 'a:xfrm', @doc
    offset =  Nokogiri::XML::Node.new 'a:off', @doc
    offset['x'] = '0'
    offset['y'] = '0'
    transform.add_child offset
    extents =  Nokogiri::XML::Node.new 'a:ext', @doc

    extents['cx'] = (emu_dimensions[0] * 1.0000762).round
    extents['cy'] = (emu_dimensions[1] * 1.0000762).round
    transform.add_child extents
    shape_properties.add_child transform

    preset_geometry =  Nokogiri::XML::Node.new 'a:prstGeom', @doc
    preset_geometry['prst'] = 'rect'
    adjust_value_list = Nokogiri::XML::Node.new 'a:avLst', @doc
    preset_geometry.add_child adjust_value_list
    shape_properties.add_child preset_geometry

    no_fill = Nokogiri::XML::Node.new 'a:noFill', @doc
    shape_properties.add_child no_fill

    line = Nokogiri::XML::Node.new 'a:ln', @doc
    line['w'] = 0 # '9525'
    line.add_child no_fill.dup
    miter = Nokogiri::XML::Node.new 'a:miter', @doc
    miter['lim'] = '800000'
    line.add_child miter
    line.add_child Nokogiri::XML::Node.new 'a:headEnd', @doc
    line.add_child Nokogiri::XML::Node.new 'a:tailEnd', @doc
    shape_properties.add_child line

    pic.add_child shape_properties

    drawing
  end

  def process_runs(source_node, cur_run)
    old_run = cur_run

    source_node.children.each do |run|
      next_run = cur_run.dup
      if cur_run.parent.nil?
        old_run.add_next_sibling(cur_run)
      end
      if run.name == 'strong'
        bold = Nokogiri::XML::Node.new 'w:b', @doc
        get_node(cur_run, 'w:rPr').add_child bold
      elsif run.name == 'br'
        br = Nokogiri::XML::Node.new 'w:br', @doc
        get_node(cur_run, 'w:rPr').add_child br
      elsif run.name == 'a'
        rel_id = add_rels_link(run['href'], :hyperlink)
        r_id = "rId#{rel_id}"
        hyperlink = Nokogiri::XML::Node.new 'w:hyperlink', @doc
        hyperlink['r:id'] = r_id
        hyperlink['w:history'] = "1"
        cur_run.add_next_sibling(hyperlink)
        hyperlink.add_child(cur_run)
        style = Nokogiri::XML::Node.new 'w:rStyle', @doc
        style['w:val'] = 'Hyperlink'
        get_node(cur_run, 'w:rPr').add_child style
      elsif run.name == 'img'
        image_node = create_image_node run['src'], run['alt']
        run_properties = get_node(cur_run, 'w:rPr')
        run_properties.add_child Nokogiri::XML::Node.new('w:noProof', @doc)
        #(cur_run.parent/'.//w:ind').remove
        lang = Nokogiri::XML::Node.new 'w:lang', @doc
        lang['w:val'] = 'en-GB'
        lang['w:eastAsia'] = 'en-GB'
        run_properties.add_child lang
        cur_run.add_child image_node
      end
      if run.name != 'img'
        if run.text[0] == ' '
          new_text = Nokogiri::XML::Node.new "w:t", @doc
          cur_run.add_child new_text
          new_text['xml:space'] = "preserve"
          new_text.content = ' '
          old_run = cur_run
          cur_run = next_run
          next_run = cur_run.dup
          old_run.add_next_sibling(cur_run)
        end

        new_text = Nokogiri::XML::Node.new "w:t", @doc
        cur_run.add_child new_text
        new_text.content = run.text

        if run.text[run.text.length-1] == ' '
          cur_run.add_next_sibling(next_run)
          cur_run = next_run
          next_run = cur_run.dup
          new_text = Nokogiri::XML::Node.new "w:t", @doc
          cur_run.add_child new_text
          new_text['xml:space'] = "preserve"
          new_text.content = ' '
        end
      end

      if run.name == 'a'
        old_run = hyperlink
      else
        old_run = cur_run
      end
      cur_run = next_run
    end
  end

  def add_paragraph(text_elements, cur_node)
    first_run = (cur_node.xpath('.//w:r')).first

    process_runs text_elements, first_run
  end

  def add_bullet(text_elements, cur_node)
    bullet_node = (cur_node.xpath('.//w:pPr')).first
    (bullet_node/'.//w:ind').remove
    numpr = Nokogiri::XML::Node.new "w:numPr", @doc
    ilvl = Nokogiri::XML::Node.new "w:ilvl", @doc
    ilvl['w:val'] = '0'
    numpr.add_child ilvl
    numId = Nokogiri::XML::Node.new 'w:numId', @doc
    numId['w:val'] = '22'
    numpr.add_child numId
    bullet_node.add_child numpr

    first_run = (cur_node.xpath('.//w:r')).first
    process_runs text_elements, first_run
  end

  def process_elements(elements, cur_node)
    old_node = cur_node
    elements.each do |element|
      dup_p = cur_node.dup
      if cur_node.parent.nil?
        old_node.add_next_sibling(cur_node)
      end
      if element.name == 'p'
        add_paragraph element, cur_node
      elsif element.name == 'h3'
        properties = get_node(cur_node, 'w:pPr')
        style = get_node(properties, 'w:pStyle')
        style['w:val'] = 'Heading3'
        add_paragraph element, cur_node
      elsif element.name == 'li'
        add_bullet element, cur_node
      elsif element.name == 'ul'
        cur_node = process_elements element.elements, cur_node
      end

      old_node = cur_node
      cur_node = dup_p
    end
    old_node
  end
  
  def merge_yaml(yaml_file)
    records = YAML::parse_file yaml_file
    merge(records)
  end

  def merge(rec)
    xml = @zip.read("word/document.xml")
    @doc = Nokogiri::XML(xml) {|x| x.noent}
    doc_values = (@doc/"//w:p")  #.select { |element| element.text =~ /^\$(.*)\$$/  }
    doc_values.each do |element|
      #puts element.text
      if element.text =~ /^\$(.*)\$$/
        value = (rec.select $1)[0]
        if value.nil?
          value = ''
        else
          value = value.value
        end

        # Ensure there is only one w:t node
        (element/'.//w:t').remove

        value_html = Kramdown::Document.new(value).to_html

        html_doc = Nokogiri::HTML(value_html)
        if html_doc.elements.count > 0
          markdown_elements = html_doc.elements[0].elements[0].elements
          process_elements(markdown_elements, element)
        end
      end
    end
    @replace["word/document.xml"] = @doc.serialize :save_with => 0
  end

  def check_for_jpeg_content_type()
    xml_file = '[Content_Types].xml'
    xml = @replace[xml_file] || @zip.read(xml_file)
    types = Nokogiri::XML(xml) {|x| x.noent}

    default_types = types.elements[0]/"Default"

    if default_types.select { |element| element['Extension'] == 'jpeg'}.count == 0
      jpeg_type = Nokogiri::XML::Node.new 'Default', types
      jpeg_type['Extension'] = 'jpeg'
      jpeg_type['ContentType'] = 'image/jeg'
      #default_types.last.add_next_sibling jpeg_type
      types.children[0].add_child jpeg_type
      @replace[xml_file] = types.serialize :save_with => 0
    end


  end

  def add_rels_link(url, type)
    xml = @replace["word/_rels/document.xml.rels"] || @zip.read("word/_rels/document.xml.rels")
    rels = Nokogiri::XML(xml) {|x| x.noent}

    cur_id = 1
    have_unused_id = false
    until have_unused_id
      id = "rId#{cur_id}"
      have_unused_id = true
      rels.elements[0].elements.each do |element|
        if element['Id'] == id
          cur_id += 1
          id = "rId#{cur_id}"
          have_unused_id = false
        end
      end
    end

    relationship = Nokogiri::XML::Node.new "Relationship", rels
    relationship['Id'] = id
    if type == :hyperlink
      relationship['Type'] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
      relationship['TargetMode'] = 'External'
    elsif type == :image
      relationship['Type'] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    end
    relationship['Target'] = url

    rels.elements[0].add_child relationship

    @replace["word/_rels/document.xml.rels"] = rels.serialize :save_with => 0

    cur_id
  end

  def save(path)
    Zip::File.open(path, Zip::File::CREATE) do |out|
      @zip.each do |entry|
        out.get_output_stream(entry.name) do |o|
          if @replace[entry.name]
            o.write(@replace[entry.name])
          else
            o.write(@zip.read(entry.name))
          end
        end
      end
      @media.keys.each do |key|
        out.get_output_stream(key) { |o| o.write @media[key] }
      end
    end
    @zip.close

    # this is to ensure the zip can actually be opened by word. RubyZip doesn't quite do it
    # from the build server
    require 'rbconfig'
    is_windows = (RbConfig::CONFIG['host_os'] =~ /mswin|mingw|cygwin/)
    if is_windows
      if Dir.exist?('releasePack')
        FileUtils.rmtree 'releasePack'
      end
      `"C:\\program files\\7-zip\\7z.exe" x -oreleasePack release.docx`
      `cd releasePack && "C:\\program files\\7-zip\\7z.exe" a release.docx *`
      FileUtils.move File.join('releasePack', 'release.docx'), 'release.docx'
      FileUtils.rmtree 'releasePack'
    end
  end
end

