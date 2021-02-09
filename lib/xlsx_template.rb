require "xlsx_template/version"

module XlsxTemplate
  class Error < StandardError; end

  def self.load_file(path, before_row_column_index: 0)
    Template.load_file(path, before_row_column_index: before_row_column_index)
  end
end

require "xlsx_template/template"
