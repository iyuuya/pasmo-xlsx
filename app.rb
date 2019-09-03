# frozen_string_literal: true

$LOAD_PATH.unshift File.join(__dir__, 'lib')

APP_ENV = ENV.fetch('APP_ENV', 'production').to_sym

require 'bundler/setup'
Bundler.require(:default, APP_ENV)

configure :development do
  register Sinatra::Reloader
end

require 'nkf'
require 'csv'
require 'pasmo_csv_converter'

set :environment, APP_ENV

get '/' do
  erb :index
end

post '/upload' do
  PasmoCsvConverter.create_temp_directory
  PasmoCsvConverter.remove_temp_files

  converter = PasmoCsvConverter.new(
    params[:csv][:tempfile].read,
    params[:csv][:filename]
  )
  converter.call

  send_file converter.xlsx_path, filename: converter.xlsx_name
end
