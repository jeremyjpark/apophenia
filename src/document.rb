# typed: true
require "sablon"

class Document
  def initialize(path_to_file:)
    @parser = Sablon::Parser::MailMerge.new
    @file_path = path_to_file
  end

  def tags_in_document
    template_file_name = File.expand_path(@file_path)
    template = Sablon.template(template_file_name)

      untouched_template = Sablon::DOM::Model.new(Zip::File.open(@file_path))
      fields = []

      files = untouched_template.zip_contents.keys.select { |key|
        key.start_with?(/word\/document.xml/, /\word\/footer.*.xml/, /word\/header.*.xml/)
      }

      files.each do |file|
        fields << @parser.parse_fields(untouched_template.zip_contents[file]).map(&:expression)
      end

      # figure out ways to identify different types of fields.
      fields.flatten.uniq
  end

  def cleaned_tags
    tags_in_document.map { |tag| tag.gsub(/^=/, "") }
  end

  def render_with_tags
    template_file_name = File.expand_path(@file_path)
    template = Sablon.template(template_file_name)
    output_file = File.join(__dir__, "rendered.docx")

    template.render_to_file(output_file, fields)
  end

  def fields
    {
      locations: {
        location_1: {
          detail1: "foo",
          detail2: "detail2",
          detail3: "detail3"
        },
        location_2: {
          detail1: "bar",
          detail2: "detail2",
          detail3: "detail3"
        }
      },
      forms_list: [
      ]
    }

    # {
    #   locations: [
    #     {
    #       index: "1",
    #       detail1: "foo",
    #       detail2: "detail2",
    #       detail3: "detail3"
    #     },
    #     {
    #       index: "2",
    #       detail1: "bar",
    #       detail2: "detail2",
    #       detail3: "detail3"
    #     }
    #   ],
    # }

  end
end