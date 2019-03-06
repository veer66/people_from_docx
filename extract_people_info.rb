require 'zip'
require 'pathname'
require 'docx'
require 'json'

$year = '2019'

def extract_pic(path)
  Zip::File.open(path) do |zip|
    zip.each do |entry|
      extname = Pathname.new(entry.name).extname
      if extname =~ /.jpeg|.png/
        content = entry.get_input_stream.read
        return {content: content, extname: extname}
      end
    end
  end
  return nil
end

def extract_person_info(path)
  pn = Pathname.new(path)
  title = pn.basename.to_s.sub(/\.docx$/,'')
  filename = title.gsub(/[,\s\.]+/, '_').downcase
  pic = extract_pic(path)
  if pic
    pic_path = filename + pic[:extname]
    File.open(File.join("pic", pic_path), 'w') do |file|
      file.write(pic[:content])
    end
  else
    pic_path = "nopic.png"
  end
  doc = Docx::Document.open(path)
  doc = doc.paragraphs.map{|p| p.text}.select{|p| p =~ /[^\s]+/}
  name = doc[0]
  countries = doc[1].split(/,/).map{|country| country.chomp}
  content = doc[2..-1]
  md_content = content.join("\n\n")
  md_path = "#{$year}_#{filename}.md"
  File.open(File.join("data", md_path), 'w') do |file|
    file.write(md_content)
  end
  
  {pic: pic_path,
   bio: md_path,
   name: name,
   countries: countries}
end

def main
  people = []
  Dir::glob("source/*.docx") {|path|
    person_info = extract_person_info(path)
    people << person_info
  }
  info = {year: $year, people: people}
  File.open(File.join("data", "award#{$year}.json"), 'w') do |file|
    file.write(JSON.pretty_generate(info))
  end
end

main if $0 == __FILE__


                        



                        
