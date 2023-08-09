#!/usr/bin/env ruby

# Convert FLP Leaves shelf marks to folder names

def normalize string
  return '' unless string
  string.strip.downcase.gsub(/\s+/, '_')
end

ARGF.each do |line|
  text = line.strip
  case text
  when /^Call Number\/ID/
    puts text
  when /^(Lewis E M )(\d+)[:.](\d+)([a-zA-Z]?)-(\d+)([a-zA-Z]?)$/
    conv = sprintf "%s%03d %03d%s-%03d%s", $1, $2, $3, $4, $5, $6
    puts normalize conv
  when /^(Lewis E M )(\d+)[:.](\d+)(.*)$/
    conv = sprintf "%s%03d %03d%s", $1, $2, $3, $4
    puts normalize conv
  when /^(Lewis T)(\d+)(.*)$/
    conv = sprintf"%s%03d%s", $1, $2, $3
    puts normalize conv
  when /^(Lewis Add )(\d+)(.*)$/
    conv = sprintf"%s%03d%s", $1, $2, $3
    puts normalize conv
  else
    puts "NOT_HANDLED: #{text}"
  end

end
