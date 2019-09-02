# frozen_string_literal: true

require 'nkf'
require 'csv'
require 'axlsx'

require 'sinatra'

get '/' do
  erb :index
end

post '/upload' do
  Dir.mkdir('tmp') unless Dir.exist? 'tmp'
  Dir['tmp/*.xlsx'].each { |f| File.delete f }

  file = params[:csv][:tempfile]
  basename = File.basename(params[:csv][:filename], '.csv')
  xlsx_name = "#{basename}.xlsx"
  path = "tmp/#{xlsx_name}"

  Axlsx::Package.new do |package|
    package.use_autowidth = true

    package.workbook do |wb|
      borders = wb.styles.add_style border: { style: :thin, color: 'F000000', name: :borders, edges: %i[top right bottom left] }

      wb.add_worksheet(name: basename, page_setup: { fit_to_width: 1, orientation: :landscape, paper_size: 9 }) do |sheet|
        count = 0
        CSV.parse(NKF.nkf('-w', file.read)) do |row|
          if count == 0
            sheet.add_row row
          else
            sheet.add_row row, style: borders
          end
          count += 1
        end
        sheet.add_row [nil, nil, nil, nil, nil, nil, '合計金額', "=SUMIFS(H3:H#{count}, H3:H#{count}, \">0\")", nil, nil], style: borders
      end
      package.serialize path
    end
  end

  send_file path, filename: xlsx_name
end
