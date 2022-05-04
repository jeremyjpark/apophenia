# typed: false
require "sinatra"
require "./src/document"

get "/" do
  document = Document.new(path_to_file: "docs/location_loops.docx")
  tags = document.tags_in_document
  document.render_with_tags
  puts tags
  puts document.cleaned_tags
end
