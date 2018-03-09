#!/usr/bin/env ruby

# Add the lib dir to our load path
$LOAD_PATH.unshift(File.expand_path '../lib', __FILE__)

require 'pp'
require 'rubyXL'
require 'genizah'

include Genizah::Util

# We want cell D8
ADMIN_EMAIL_ROW = 7
ADMIN_EMAIL_COL = 3

xlsx_input = ARGV.shift
fail "Argument is not an XLSX file '#{xlsx_input}'" unless xlsx_file? xlsx_input

workbook = RubyXL::Parser.parse xlsx_input
worksheet = workbook['Description']

cell = get_cell worksheet, ADMIN_EMAIL_ROW, ADMIN_EMAIL_COL
puts cell.inspect
value = cell.value
puts "Value is #{value}"
cell.change_contents value
cell.style_index = 0
puts "After updates #{cell.inspect}"

workbook.write xlsx_input
puts "Rewrote '#{xlsx_input}'."
