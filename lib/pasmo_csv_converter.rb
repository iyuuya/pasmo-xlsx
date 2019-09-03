# frozen_string_literal: true

require 'nkf'
require 'csv'
require 'axlsx'

class PasmoCsvConverter
  def initialize(csv_data, csv_filename)
    @csv_data = csv_data
    @csv_filename = csv_filename
    @basename = File.basename(@csv_filename, '.csv')
  end

  def call
    Axlsx::Package.new do |package|
      package.use_autowidth = true

      package.workbook do |wb|
        borders = wb.styles.add_style border: {
          style: :thin,
          color: 'F000000',
          name: :borders,
          edges: %i[top right bottom left]
        }

        wb.add_worksheet(name: @basename, page_setup: { fit_to_width: 1, orientation: :landscape, paper_size: 9 }) do |sheet|
          count = 0
          CSV.parse(NKF.nkf('-w', @csv_data)) do |row|
            if count == 0
              sheet.add_row row
            else
              sheet.add_row row, style: borders
            end
            count += 1
          end

          row = sheet.add_row [nil, nil, nil, nil, nil, nil, '合計金額', "=SUMIFS(H3:H#{count}, H3:H#{count}, \">0\")", nil, nil]
          row.cells.to_ary.select { |cell| !cell.value.nil? }.each { cell.style = borders }
        end
        package.serialize xlsx_path
      end
    end
  end

  def xlsx_path
    @xlsx_path ||= "tmp/#{xlsx_name}"
  end

  def xlsx_name
    @xlsx_name ||= "#{@basename}.xlsx"
  end

  class << self
    def create_temp_directory
      Dir.mkdir('tmp') unless Dir.exist? 'tmp'
    end

    def remove_temp_files
      Dir['tmp/*.xlsx'].each { |f| File.delete f }
    end
  end
end
