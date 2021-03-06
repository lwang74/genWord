#coding UTF-8
Encoding.default_internal = 'UTF-8'

# 教育学会个人会员注册_20151203

require 'fileutils'
require_relative 'excel'

class Word
	def initialize doc_file
		@curr_path = FileUtils.pwd.gsub(/\//, '\\').encode('UTF-8')
		@word_app = WIN32OLE.new('Word.Application')
		@doc = @word_app.Documents.open("#{@curr_path}\\#{doc_file}", 'ReadOnly' => true)
		@tables = @doc.tables
	end

	def write_in_table text, table, row, column
		@tables.Item(table).cell(row, column).range.text = text
	end

	def add_picture_in_table pic_path, table, row, column
		full_pic_path = "#{@curr_path}\\#{pic_path}\\*.jpg"
		pics = Dir[full_pic_path.gsub(/\\/, '/').encode(Encoding.default_external)]

		@pic = true
		if pics.size==1
			# puts pics[0]
			shp = @tables.Item(table).cell(row, column).range.inlineShapes.AddPicture(pics[0])
			shp.height = @word_app.CentimetersToPoints(4) 
		else
			@pic = false
			# puts "找不到照片 #{pic_path} !!!"
		end
	end

	def save_as target_file
		no_pic = @pic? '' : ' (无照片)'
		@doc.SaveAs "#{@curr_path}\\#{target_file} #{no_pic}.doc"
		@word_app.quit
	end
end

def main
	begin
	out = "output"
	# FileUtils.rm_rf out
	FileUtils.mkdir out if !File.exist?(out)

	excel = '河东区八十二中会员花名册.xls'
	sht = '花名册'
	word_file = "附件1.会员证模板（2016年）.doc"

	all_list = []
	CExcel.new.open_read(excel.encode(Encoding.default_external)){|wb|
		wb.Worksheets(sht).usedrange.value2.each{|row|
			all_list<<row if row.compact.size>0
		}
	}

	all_info = all_list.map{|row|
		# p row[9]
		[row[0], row[1], row[2], row[3], row[4], row[5]] if /^HY\-\d{3}\-z\d{3}\-\d{3}$/=~ row[0] and (row[9]!=nil and row[9].to_s=~/\S+/)
	}.compact

	# all_info.each{|row|
	# 	puts row[1]
	# }

	# delete exists Word file
	deleting_file = nil
	all_info.each{|row|
		# puts "#{out}\\#{row[0]}+*.doc"
		match_files = Dir["#{out}/#{row[0]}+*.doc"]
		match_files.each{|one|
			deleting_file = one
			FileUtils.rm one
		}
	}

	cnt = 1
	all_info.each{|one|
		print '.'
		word = Word.new(word_file)
		word.write_in_table one[0], 1, 8, 4
		word.write_in_table one[1], 1, 1, 2
		word.write_in_table one[2], 1, 2, 2
		word.write_in_table one[3], 1, 3, 2
		word.write_in_table "#{one[4]}, #{one[5]}", 1, 5, 2
		code = one[0].encode('UTF-8')
		name = one[1].encode('UTF-8')
		pic = "photos\\#{code}"
		word.add_picture_in_table pic, 1, 1, 3	
		word.save_as "#{out}\\#{one[0]}+#{one[1]}"
		# break if cnt>=20
		cnt+=1
	}
	rescue Errno::EEXIST=> e
		puts "路径'#{out}' 被占用，请手动删除后重试！"
	rescue Errno::EACCES=> e
		puts "Word文件'#{deleting_file}' 被占用，请关闭相应Word文件后重试！"
	end
end

main
