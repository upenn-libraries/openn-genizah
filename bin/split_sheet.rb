#!/usr/bin/env ruby

# Add the lib dir to our load path
$LOAD_PATH.unshift(File.expand_path '../../lib', __FILE__)

require 'yaml'
require 'pp'
require 'rubyXL'
require 'rubyXL/convenience_methods/cell'
require 'genizah'

include Genizah::Util

COLUMN_MAP = Hash.new { |hash,key|
  hash[key] = hash[key - 1].next
}.merge({0 => "A"})

TEMPLATE_HEADING_COLUMN = 1 # Excel col: B
VALUE_START_COLUMN      = 3 # Excel col: D
DISPLAY_PAGE_COLUMN     = 1 # Excel col: B
FILE_NAME_COLUMN        = 2 # Excel col: C
PAGES_START_ROW         = 3 # Excel row: 4

MAPPING = YAML::load open(File.expand_path '../../mapping.yml', __FILE__)
# SOURCE_DIR = File.expand_path '../data/FLP', __FILE__
# SOURCE_DIR = '/Volumes/mmscratchspace/openn/packages/Prep/genizah'
# SOURCE_DIR = '/mnt/scratch02/openn/packages/Prep/flp_leaves'
SOURCE_DIR = '/Users/emeryr/tmp/geniza_files/data'

# IDS_FILE = File.expand_path '../halper_ids.txt', __FILE__

CREATE_DIRS    = ENV['GENIZAH_CREATE_DIRS']    || false
ALLOW_NO_TIFFS = ENV['GENIZAH_ALLOW_NO_TIFFS'] || false
NO_CLOBBER     = ENV['GENIZAH_NO_CLOBBER']     || false
REQUIRE_TIFFS  = ENV['REQUIRE_TIFFS']          || true

# HalperCallNumber = Struct.new :folder, :mark, :url
# HALPER_IDS = open(IDS_FILE).readlines.inject({}) { |h, line|
#   callno = HalperCallNumber.new *line.strip.split(/\t/)
#   h[callno.folder] = callno unless callno.url == 'NONE'
#   h
# }

RV = %w{ r v }

#----------------------------------------------------------------------
# METHODS
#----------------------------------------------------------------------

def mapped? heading
  MAPPING.include? heading.to_sym
end

def cell_empty? cell
  cell.nil? || cell.value.nil? || cell.value.to_s.strip.empty?
end

# def xlsx_file? path
#   return false if path.nil? || path.to_s.strip.empty?
#   File.exists?(path) && path =~ /\.xlsx?/
# end

def find_row worksheet, heading, col=TEMPLATE_HEADING_COLUMN
  blank_count = 0
  (0..200).each do |i|
    break if blank_count > 5
    cell = worksheet[i][col]
    cell_empty? cell and blank_count += 1 and next
    blank_count = 0
    return i if normal_head(cell.value) == heading.to_s
  end
end

def normal_head value
  value.downcase.strip.gsub /[^[:alnum:]]+/, '_'
end

# def get_cell sheet, row, col, default=''
#   return sheet.add_cell row, col, default if sheet[row].nil?
#   return sheet.add_cell row, col, default if sheet[row][col].nil?
#   sheet[row][col]
# end

def transform val, rule
  case rule
  when 'YEAR'
    raise "Not a valid integer: #{val}" unless is_i? val
    sprintf "%04d", val.to_i
  when 'REMOVE_TRAILING_PERIOD'
    # remove a final '.'
    # for cleaning up keyword terms
    val.sub /\.\s*$/, ''
  when 'STRING_CLEANUP'
    # replace a sequence '.;' or '. ;' with '. ': '...blah blah.; Yadda ..' becomes '...blah blah. Yadda ..''
    # proper spacing around semicolons 'text;text' as 'text; text'
    # remove a trailing ':'
    val.gsub(/\.\s*;/, '. ').gsub(/\s*;\s*/, '; ').sub(/\s*:\s*$/, '').strip
  else
    val
  end
end

def is_i? val
 /\A[-+]?\d+\z/ === val.to_s
end

#----------------------------------------------------------------------
# INPUT
#----------------------------------------------------------------------

xlsx_input = ARGV.shift
fail "Argument is not an XLSX file '#{xlsx_input}'" unless xlsx_file? xlsx_input

#----------------------------------------------------------------------
# CONFIGURE
#----------------------------------------------------------------------
template_path = File.expand_path '../../data/template.xlsx', __FILE__
template = RubyXL::Parser.parse template_path
# description = template['Description']

MAPPING.each do |heading, deets|
  row = find_row template['Description'], heading
  raise "Could not find heading: #{heading}" if row.nil?
  deets[:row] = row
end

workbook = RubyXL::Parser.parse xlsx_input
worksheet = workbook[0]

headers = worksheet[0].cells.map do |cell|
  next if cell.nil?
  normal_head(cell.value).to_sym
end

folder_base_index = headers.index :folder_base

#----------------------------------------------------------------------
# RUN
#----------------------------------------------------------------------

(1..2000).each do |rowindex|
  break if worksheet[rowindex].nil?
  row = worksheet[rowindex]

  begin
    # -- The directory
    folder_base = row[folder_base_index].value
    if folder_base.nil? or folder_base.strip.empty?
      address = "#{COLUMN_MAP[folder_base_index]}#{rowindex}"
      raise "Row doesn't have folder_base expected at #{address}"
    end
    folder = File.join SOURCE_DIR, folder_base
    # binding.pry
    if Dir.exists? folder
      if Dir["#{folder}/*.tif"].empty? && REQUIRE_TIFFS
        STDERR.puts "Skipping folder without TIFFs: '#{folder}'; row #{rowindex}"
        next
      end
    elsif CREATE_DIRS
      Dir.mkdir folder
    else
      STDERR.puts "Skipping folder creation, row #{rowindex}, folder_base '#{folder_base}'"
      next
    end


    out_xlsx = File.join(folder, "openn_metadata.xlsx")
    if File.exists? out_xlsx
      if NO_CLOBBER
        STDERR.puts "Skipping existing file '#{out_xlsx}'"
        next
      else
        STDERR.puts "Overwriting existing file '#{out_xlsx}'"
      end
    end
    FileUtils.cp template_path, out_xlsx
    outbook = RubyXL::Parser.parse out_xlsx

    description = outbook['Description']
    headers.each_with_index do |head, hindex|
      next unless mapped? head
      cell = row[hindex]
      deets = MAPPING[head]
      next if cell_empty? cell
      cell.value.to_s.split(/\|/).each_with_index do |val, j|
        target_row = deets[:row]
        target_col = VALUE_START_COLUMN + j
        cell = get_cell description, target_row, target_col
        cell.change_contents transform(val, deets[:transform])
      end
    end

    pages = outbook['Pages']

    # binding.pry
    if headers.index :image_files
     files  = row[headers.index :image_files].value.to_s.strip.chomp('|').split('|')
     labels = row[headers.index :image_labels].value.to_s.strip.chomp('|').split('|')
     unless files.size == labels.size
      STDERR.puts "WARNING: files do not match labels: #{files}/#{labels}"
      STDERR.puts "SKIPPING: row: #{rowindex + 1} -- #{row[headers.index :call_number_id].value}"
      next unless File.exist? out_xlsx
      STDERR.puts "Removing output file '#{out_xlsx}'" if File.exist? out_xlsx
      FileUtils.rm out_xlsx if File.exist? out_xlsx
      next
    end

    file_labels = files.zip labels
    file_labels.each_with_index do |file_label, k|
      file, label = file_label
      base = File.basename file
      file_cell = get_cell pages, PAGES_START_ROW + k, FILE_NAME_COLUMN
      file_cell.change_contents base
      page_cell = get_cell pages, PAGES_START_ROW + k, DISPLAY_PAGE_COLUMN
      page_cell.change_contents label
    end
  else
    Dir["#{folder}/*.tif"].each_with_index do |tif,k|
      base = File.basename tif
      file_cell = get_cell pages, PAGES_START_ROW + k, FILE_NAME_COLUMN
      file_cell.change_contents base
      page_cell = get_cell pages, PAGES_START_ROW + k, DISPLAY_PAGE_COLUMN
      page_num = "#{(k + 2)/ 2}#{RV[k % 2]}"
      page_cell.change_contents page_num
    end
  end

    outbook.write out_xlsx
    puts "Wrote #{out_xlsx}"
  rescue
    STDERR.puts $!
    STDERR.puts $!.backtrace
    STDERR.puts "Error processing row #{rowindex} and folder_base '#{folder_base}'"
    STDERR.puts "Removing output file '#{out_xlsx}'"
    # binding.pry
    FileUtils.rm out_xlsx if out_xlsx && File.exist?(out_xlsx)
  end
end
