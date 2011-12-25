class IrasParser

  require "nokogiri"

  def initialize(xml_filename)
    @xml = xml_filename
    @doc = ""
    @xml_map = []

    File.open(@xml) do |file|

      # We need to remove all extra white spaces and new lines
      # so that Nokogiri can process the XML properly

      xml_text = ""
      file.each_line do |line|
        xml_text = xml_text + line.strip
      end
      @doc = Nokogiri::XML(xml_text)
    end
  end

  # Parses the XML file and store the values in a array of key value pairs
  # returns Array

  def process_xml
    @doc.css("Details").children.each do |ir8arecord|
      data_map = {}

      # IR8ARecord -->  ESubmissionSDSC --> IR8AST

      ir8arecord.child.child.children.each do |data|
        data_map[data.name] = data.content
      end

      @xml_map << data_map
    end

    @xml_map

  end

end

class IrasExcel
  require 'spreadsheet'

  def initialize(iras_data)

    @spreadsheet = Spreadsheet::Workbook.new
    @worksheet = @spreadsheet.create_worksheet :name => "IRAS"

    @row = 0

    prepare(iras_data)
  end

  def prepare(iras_data)
    insert_header(iras_data)
    insert_data(iras_data)
  end

  def write
    @spreadsheet.write "output.xls"
  end

  private

  # Insert the headers
  def insert_header(iras_data)
    @worksheet.insert_row @row, iras_data.first.keys
    next_row
  end

  def insert_data(iras_data)
    iras_data.each do |data|
      @worksheet.insert_row @row, data.values
      next_row
    end
  end

  def skip_row(num)
    @row += num
  end

  def next_row
    skip_row(1)
  end



end

iras = IrasParser.new("#{ARGV[0]}")
iras_data = iras.process_xml
IrasExcel.new(iras_data).write
