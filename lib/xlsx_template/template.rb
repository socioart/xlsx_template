require "rubyXL"
require "rubyXL/convenience_methods/worksheet"
require "rubyXL/convenience_methods/cell"
require "ruby-handlebars"
require "cgi"

module XlsxTemplate
  class Template
    attr_reader :workbook, :before_row_column_index, :handlebars

    def self.load_file(path, before_row_column_index: 0)
      new(RubyXL::Parser.parse(path), before_row_column_index: before_row_column_index)
    end

    def initialize(workbook, before_row_column_index: 0)
      @workbook = workbook
      @before_row_column_index = before_row_column_index
      @handlebars = ::Handlebars::Handlebars.new
    end

    def render(path, variables)
      workbook.worksheets.each do |worksheet|
        next unless worksheet[0]
        next unless worksheet[0][before_row_column_index]&.value == "beforeRow"

        worksheet[0][before_row_column_index].change_contents("") # beforeRow の値を出力に含めない

        row_index = 1 # 処理中に行数が変わるので each など使わず index で
        loop do
          break unless row_index < worksheet.sheet_data.rows.size

          row = worksheet[row_index]
          if row.nil?
            row_index += 1
            next
          end

          before_row = parse_before_row(row[before_row_column_index]&.value)
          worksheet[row_index][before_row_column_index].change_contents("") if before_row # beforeRow の値を出力に含めない

          case before_row&.first
          when :if
            variable = variables.fetch(before_row[1].to_sym)
            unless truthy?(variable)
              worksheet.delete_row(row_index)
              next # 削除したので row_index 増やさず次へ
            end
          when :each
            template_row_index = row_index
            items = variables.fetch(before_row[1].to_sym)

            unless items.empty?
              row_height = worksheet.get_row_height(row_index)
              values = row.cells.map {|c| c&.value }
              merges = worksheet.merged_cells.select {|mc| mc.ref.row_range == (row_index..row_index) }.map {|mc| mc.ref.col_range }

              items.each do |item|
                row_index += 1
                worksheet.insert_row(row_index)
                worksheet.change_row_height(row_index, row_height)
                merges.each do |m|
                  worksheet.merge_cells(row_index, m.first, row_index, m.last)
                end

                fill_values(worksheet, row_index, values)
                compile_cells(worksheet, row_index, variables.merge(item))
              end
            end

            worksheet.delete_row(template_row_index)
            next # 削除したので row_index 増やさず次へ
          end

          compile_cells(worksheet, row_index, variables)
          row_index += 1
        end

        # delete_column だとデータが壊れるので非表示
        worksheet.cols[before_row_column_index].hidden = true
      end

      workbook.write(path)
    end

    private
    def parse_before_row(s)
      case s.to_s.strip
      when ""
        nil
      when /\A\{\{#if (.+)\}\}\z/
        [:if, $~.captures[0]]
      when /\A\{\{#each (.+)\}\}\z/
        [:each, $~.captures[0]]
      else
        raise "Unknown beforeRow value #{s.inspect}"
      end
    end

    def compile_cells(worksheet, row_index, stackframe)
      row = worksheet[row_index]

      ((before_row_column_index + 1)...row.cells.size).each do |column_index|
        cell = worksheet[row_index][column_index]
        next if cell.nil?
        next if cell.value.nil?

        case cell.datatype
        when "s", "str"
          # noop
        else
          next
        end

        compiled = handlebars.compile(cell.value).call(stackframe)

        worksheet[row_index][column_index].change_contents(compiled)
      end
    end

    def fill_values(worksheet, row_index, values)
      row = worksheet[row_index]
      ((before_row_column_index + 1)...row.cells.size).each do |column_index|
        cell = worksheet[row_index][column_index]
        value = values[column_index]
        next if cell.nil?
        next if value.nil?

        worksheet[row_index][column_index].change_contents(value)
      end
    end

    # https://handlebarsjs.com/guide/builtin-helpers.html#if
    def truthy?(v)
      case v
      when nil, false, "", 0, []
        false
      else
        true
      end
    end
  end
end
