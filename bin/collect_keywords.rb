#!/usr/bin/env ruby

require 'csv'

def output_array row_in
  return [row_in['Call Number/ID'], ''] unless row_in['Facets']
  [row_in['Call Number/ID'], row_in['Facets'].split(/,\s+/).join('|')]
end

CSV do |csv|
  csv << %w{ call_number_id facets }

  CSV.foreach ARGV.first, headers: true do |row|
    begin
      csv << output_array(row) if $..odd?
    rescue
      STDERR.puts "ERROR: #{row.to_s.strip}; #{$!.message}"
    end
  end
end
