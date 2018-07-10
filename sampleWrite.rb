#RubyでExcelファイルを操作するプログラム

require "rubygems"
require "spreadsheet"

#Excelファイルをインスタンス化
book = Spreadsheet::Workbook.new

#新規sheetを作成する
sheet = book.create_worksheet

#sheetの名前を設定する。
sheet.name = "Sheet1"


10.times do |n|
    sheet[n, 0] = n
    sheet[n, 1] = n * n
    sheet[n, 2] = "日本語表示"
end

#作成したbookを書き出す
book.write("sp_test.xls")
