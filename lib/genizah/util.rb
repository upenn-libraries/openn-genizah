module Genizah
  module Util
    def xlsx_file? path
      return false if path.nil? || path.to_s.strip.empty?
      File.exists?(path) && path =~ /\.xlsx?/
    end

    def get_cell sheet, row, col, default=''
      return sheet.add_cell row, col, default if sheet[row].nil?
      return sheet.add_cell row, col, default if sheet[row][col].nil?
      sheet[row][col]
    end

  end
end