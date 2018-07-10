#Excelファイルから読み込み使用するプログラム

require "rubygems"
require "spreadsheet"

#引数で指定したファイルを開く
book = Spreadsheet.open("sp_test.xls")

#sheetでbookの対象シートを指定する。
sheet = book.worksheet("Sheet1")
10.times do |n|
    printf "%3d %3d %s\n", sheet[n, 0], sheet[n, 1], sheet[n, 2]
end

