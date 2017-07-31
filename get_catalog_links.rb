#!/usr/bin/env ruby

require 'open-uri'
require 'nokogiri'

query_url = 'https://franklin.library.upenn.edu/catalog?utf8=%E2%9C%93&search_field=call_number_xfacet&q='
franklin="https://franklin.library.upenn.edu"

ARGF.each do |line|
  call_number = line.chomp
  query = "#{query_url}#{call_number.split.join('+')}"
  html = Nokogiri::HTML.parse open(query)
  node = html.xpath("//a[normalize-space(text())='#{call_number}']")[0]
  if node.nil?
    # puts "Nothing found for '#{call_number}'"
    url = "NONE"
  else
    path = node.xpath('@href').text
    url = "#{franklin}/#{path}"
  end
  folder = sprintf "h%03d", call_number.split.last.to_i
  puts sprintf( "%s\t%s\t%s", folder, call_number, url)
  sleep 1
end