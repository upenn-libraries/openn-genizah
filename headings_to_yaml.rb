#!/usr/bin/env ruby

require 'yaml'
require 'pp'

cols = "Call Number/ID
Repository Country
Repository City
Holding Institution
Repository Name
Source Collection
Date (narrative)
Date (range) start
Date (range) end
Page dimensions
Note(s)
Layout
Support Material
Language
Subject: topical
Place of origin
Script
Decoration
Author name
Former Owner name
Scribe name
Provenance details
Manuscript Name".split $/

x = cols.inject({}) { |h, col|
  normal = col.downcase.strip.gsub /[^[:alnum:]]+/, '_'
  h.merge(normal.to_sym => { head: col, transform: nil })
}

x[:date_range_start][:transform] = 'YEAR'
x[:date_range_end][:transform] = 'YEAR'
x[:provenance_details][:transform] = 'FIX_PERIOD_SEMICOLON'

puts x.to_yaml