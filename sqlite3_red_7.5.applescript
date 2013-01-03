--property sqhandler : load script file ((path to application support from user domain as string) & "Databases:sqhandler.scpt")

property folder_name : "Databases:"
property database_name : "clients"
property cr : ASCII character (13)
property LF : ASCII character (10)
property imageList : {"CI84", "CS46Signature_Code__c", "CS47Signature_Code__c", "CI69", "CI42"}
property rule_list : {"OGE", "LT", "LR", "NL", "RD", "INS", "IN1", "IN2", "CRE", "BRE"}
--copy fkname to folder_name
--return folder_name --works
---copy dbkname to database_name
--on open the_Droppings

set desktopPath to (path to desktop as Unicode text)
set appSupportPatha to path to application support from user domain
set dataFolderPatha to (appSupportPatha & folder_name) as string
tell application "Finder"
	if not (exists folder dataFolderPatha) then make folder at appSupportPatha with properties {name:(characters 1 thru -2 of folder_name as text)}
end tell --finder
set bingo to POSIX path of dataFolderPatha
set trashmelocation to quoted form of (POSIX path of desktopPath & "trashme.rb")
set sourceFilePath to quoted form of "/Users/swalton/Library/Application Support/Databases/2010bbdgrc-test.txt"
set astid to ""
set vowelcheck to {"a", "e", "i", "o", "u", "s", "A", "E", "I", "O", "U", "S"}
set thePath to "MrJr:zacrobs:"
set IDtestdoc to "Macintosh HD:Users:swalton:Documents:test outs:UA0CAP38test.indd"
set savePath to "Macintosh HD:Users:swalton:Documents:test outs:"
set Thefilename to "trickster.indd"
set thePreset to "GrizProof"
set addyBlock1 to ("[CA8Secondary ]" & return & "[CA8Primary]")
set clientID to ""
set dataList to {}
set indesignList to {}
set pkgList to {}
set ruleA to {}
set bixby_folder to "bixby outs"

----Finder parse portion
tell application "Finder"
	set savePath to (desktopPath & bixby_folder) as string
	if not (exists folder savePath) then
		make folder at desktopPath with properties {name:bixby_folder}
	end if
	set lst to selection
	set jobNo to "1239"
	set jobDesc to "thisconstant"
	set jobNo to text returned of (display dialog "What job number?" buttons {"Cancel", "OK"} default answer ("48527") default button "OK" with icon 1)
	set jobDesc to text returned of (display dialog "What pkg?" buttons {"Cancel", "OK"} default answer ("Nov test") default button "OK" with icon 1)
	repeat with thisDocumentIndex from 1 to count of lst
		if (item thisDocumentIndex of lst as string) ends with ".txt" then
			copy item thisDocumentIndex of lst to end of dataList
		else if (item thisDocumentIndex of lst as string) ends with ".indd" then
			copy item thisDocumentIndex of lst to end of indesignList
		end if --first check of dox
	end repeat --each Finder item
	--return indesignList
	if (count of dataList) = 1 then
		set text_file_PosixPath to (POSIX path of (item 1 of dataList as string))
		set text_file_path_mac to item 1 of dataList as string
	else
		beep 4
	end if -- set data file variable
	--return text_file_path_mac
end tell
---- end finder parse portion

-----Ruby portion
set myRubyscript to ("require 'csv'
require 'rubygems'
require 'sqlite3'
require 'active_record'
MY_DB_NAME = \"" & bingo as string) & "clients.db\"
MY_DB = SQLite3::Database.new(MY_DB_NAME)
ActiveRecord::Base.establish_connection(:adapter => 'sqlite3', :database => MY_DB_NAME)
class Client < ActiveRecord::Base
end
allData = File.open(ARGV[0], 'rb')
CSV::Reader.parse(allData).each_with_index do |row, index|
	if index == 0
		@column_headers = row
	end	
end
if Client.table_exists?
	ActiveRecord::Base.connection.tables.each do |y|
		ActiveRecord::Base.connection.drop_table y
	end
end
if !Client.table_exists?
  ActiveRecord::Base.connection.create_table(:clients) do |t|
  	@column_headers.each do |header|
  		t.column header, :string
  	end
  end
  puts 'hi'
end
allData = File.open(ARGV[0], 'rb')
CSV::Reader.parse(allData).each_with_index do |row, row_index|
	puts \"row #{row_index}\"
	unless row_index == 0
		client = Client.new
		@column_headers.each_with_index do |column_header, column_index|
			client.send(\"#{column_header}=\", row[column_index])
		end
		client.save
	end	
end"
--return myRubyscript
with timeout of 2400 seconds
	set fRef to (open for access file (desktopPath & "trashme.rb") with write permission)
	set eof fRef to 0
	write myRubyscript to fRef
	close access fRef
	
	--set text_file_PosixPath to quoted form of "/Users/swalton/Library/Application Support/Databases/2010bbdgrc-test.txt"
	do shell script "/usr/bin/ruby " & trashmelocation & space & quoted form of text_file_PosixPath
	tell application "Finder"
		delete (desktopPath & "trashme.rb")
	end tell
	---- end of Ruby portion
	
	
	--set text_file_path_mac to "Macintosh HD:Users:swalton:Library:Application Support:Databases:2010bbdgrc-test.txt"
	--set text_file_PosixPath to POSIX path of text_file_path
	--return text_file_PosixPath
	set file_text to read file text_file_path_mac
	set text item delimiters to LF
	set pre_file_text to text items of file_text
	set text item delimiters to astid
	set totalRecords to count of pre_file_text
	set field_names to item 1 of pre_file_text
	--return returnCommaSepQuotedString(words of field_names) --not as array
	set field_names to my tidStuff(quote, words of field_names) --seems to work
	set field_names to {"RecNo"} & field_names
	set totalFieldCt to count of field_names --seems to be right
	--return totalFieldCt
	
	-- 
	set allRun to my sql_select_all("clients")
	--return allRun --seems to work
	set allRunCT to count of allRun --count of records from sqlite
	--return allRunCT
	set text item delimiters to ASCII character 124
	set recordAll to {}
	repeat with k from 1 to allRunCT
		set recordList to {}
		set recordCT to count of every text item of item k of allRun
		repeat with m from 1 to recordCT
			set recordListBit to text item m of item k of allRun
			set end of recordList to recordListBit
		end repeat --each m of each bit of text into its own part
		set end of recordAll to recordList
		--return recordList -- record 1 in list
	end repeat --each record list
	set text item delimiters to astid
	
	-- return recordAll --seems to return correct list of lists
	--return field_names
	--for each item in recordAll
	--each item in recordAll is a client, thus must build repeat loop for each item in package
	--totalFieldCt
	set clientRecordCt to count of recordAll
	--return clientRecordCt
	tell application "Adobe InDesign CS5"
		set whole word of find change text options to false
		set include locked layers for find of find change text options to true
		set include locked stories for find of find change text options to true
		set origRange to page range of PDF export preferences
		
		
		repeat with thisClientRecord from 1 to clientRecordCt
			set thisFieldCt to count of item thisClientRecord of recordAll
			--return thisFieldCt -->17
			set imgField to {}
			set imgValue to {}
			set PosixPDF to ""
			repeat with thisPart in rule_list
				--return thisPart --> item 1 of {"OE", "LT", "LR", "NL", "RD", "INS", "IN1", "IN2", "RE"}
				repeat with thisIndesignFileIndex from 1 to count of indesignList
					tell me
						set text item delimiters to ":"
					end tell
					set thisIndesignFilePath to item thisIndesignFileIndex of indesignList as string
					set thisIndesignFile to last text item of thisIndesignFilePath as string
					tell me
						set text item delimiters to astid
					end tell
					set dotPlace to offset of "." in thisIndesignFile
					set cTchar to count of characters of thisIndesignFile
					set xEnd to characters ((cTchar) - ((cTchar) - (dotPlace - 3))) thru (cTchar - (cTchar - dotPlace + 1)) of thisIndesignFile as text
					if " " is in characters of xEnd then set xEnd to characters ((cTchar) - ((cTchar) - (dotPlace - 2))) thru (cTchar - (cTchar - dotPlace + 1)) of thisIndesignFile as text
					--return xEnd
					--return (xEnd & thisPart) --> "LROE" works
					if xEnd as string = thisPart as string then
						set iop to (xEnd & " " & thisPart)
						open (thisIndesignFilePath as alias)
						set theDoc to document 1
						set ss to ""
						tell theDoc
							set theRange to "1-" & (get count of pages of theDoc)
							set {y0, x0, pageHgt, pageWid} to bounds of page 1 of theDoc
							set thePreset to "GrizProof"
							--try
							(*
							set sampleChange to my findChangePrefs(("[DT25]" & return & "[DT94]" & return & "[DT45]" & return & "[DT33]" & return & "[DT12] [DT48] [DT61]"), ("[Mr. John Q. Sample]" & return & "[Grizzard Communications Group]" & return & "[Secondary Address]" & return & "[229 Peachtree Street, Suite 1400]" & return & "[Atlanta, GA 30303]"))
							set foundText to change text
							set simpleChange to my findChangePrefs(("[DT25]" & return & "[DT94]" & return & "[DT45]" & return & "[DT12] [DT48] [DT61]"), ("[Mr. John Q. Sample]" & return & "[Grizzard Communications Group]" & return & "[229 Peachtree Street, Suite 1400]" & return & "[Atlanta, GA 30303]"))
							set foundText to change text
*)
							--end try
							repeat with thisField from 1 to thisFieldCt
								set currentFieldName to item thisField of field_names
								set currentFieldItem to item thisField of item thisClientRecord of recordAll
								--return currentFieldItem
								--except for field 1 of item i of recordAll
								if currentFieldName is not "RecNo" then
									if currentFieldName is "CA8Secondary" and currentFieldItem is "" then
										--try
										set currentChange to my findChangePrefs(("[" & currentFieldName & "]" & return), currentFieldItem)
										set foundText to change text
										--end try
									else if currentFieldName is "CA8Primary" and "CA8Secondary" is not in field_names then
										set currentChange to my findChangePrefs(("[CA8Secondary]" & return & "[" & currentFieldName & "]" & return), currentFieldItem & return)
										set foundText to change text
									else if currentFieldName is "CA8DPBC__c" then
										-- forces "!" in -- set currentChange to my findChangePrefs(("[" & currentFieldName & "]"), "!" & currentFieldItem & "!")
										set currentChange to my findChangePrefs(("[" & currentFieldName & "]"), currentFieldItem)
										set foundText to change text
									else if currentFieldName is "Area_ID__c" then
										try
											set clientID to currentFieldItem
											set currentChange to my findChangePrefs(("[" & currentFieldName & "]"), currentFieldItem)
											set foundText to change text
										end try
										--if name is not in image then
										--find name of field in text found in same index position of field in field_names that is, for each record in recordAll use the index of field position to find index of field_names | record i, field 3 has column header of field 3 in field_names
									else if currentFieldName is in imageList then
										--try
										--set currentFieldName to end of imgField
										set imgField to imgField & currentFieldName
										--return imgField
										--set currentFieldItem to end of imgValue
										set imgValue to imgValue & currentFieldItem
										--end try
									else
										try
											set currentChange to my findChangePrefs(("[" & currentFieldName & "]"), currentFieldItem)
											set foundText to change text
										end try
										
										
									end if --check for blank secondary add
								end if --search for record non-index field
							end repeat -- each field in each clientRecord
							--return imgField
							--use image path with text frame prefs defined by item. FOR loose img frames
							set imgFieldCt to count of items in imgField
							repeat with thisImgFrame from 1 to count of rectangle
								tell rectangle thisImgFrame
									repeat with thisImgID from 1 to imgFieldCt
										if label contains (item thisImgID of imgField as text) and label contains "CI69" then
											set fitting alignment of frame fitting options to left center anchor
											--place item thisImgID of imgValue
											place ((item thisImgID of imgValue) as alias)
											redefine scaling to {1.0, 1.0}
										else if label contains (item thisImgID of imgField as text) then
											set fitting alignment of frame fitting options to top left anchor
											place ((item thisImgID of imgValue) as alias)
											redefine scaling to {1.0, 1.0}
										end if
									end repeat --each item in imgField
								end tell --each img Frame
							end repeat --each imgFrame
							--now to get image frames inside text like sig blocks.
							repeat with thisImgTextFrame from 1 to imgFieldCt
								set allRect to (every rectangle of every story whose label is (item thisImgTextFrame of imgField as text))
								
								set transform reference point of layout window 1 to bottom left anchor
								if (count of allRect) > 0 then
									repeat with thisRect in allRect
										--set ax to object reference of thisRect --of allRect
										tell thisRect
											
											set {ax} to place item thisImgTextFrame of imgValue -- on ax --with properties {fitting alignment:bottom left anchor}
											fit ax given content to frame
											set properties of ax to {absolute horizontal scale:100, absolute vertical scale:100}
										end tell --ax
									end repeat --drill down into img frames in text
								end if --allRect not empty
								set transform reference point of layout window 1 to top left anchor
								
							end repeat --second round of hunting images
						end tell --theDoc
						--close theDoc saving in (savePath & clientID & Thefilename)
						(*						set mypdfproof to my pdf81((jobNo & " " & clientID & " " & jobDesc & " " & xEnd), theDoc, theRange, thePreset, origRange, y0, x0, pageHgt, pageWid)

						copy (thePath & jobNo & " " & clientID & " " & jobDesc & " " & xEnd & ".PDF") to end of pkgList
						copy xEnd to end of ruleA
						set ss to POSIX path of (thePath & jobNo & " " & clientID & " " & jobDesc & " " & xEnd & ".PDF")
						--copy quoted form of ss to end of PosixPDF
						set PosixPDF to PosixPDF & " " & (quoted form of ss as text)
*)
						set transform reference point of layout window 1 to top left anchor
						close theDoc saving in (desktopPath & bixby_folder & ":" & jobNo & " " & clientID & " " & jobDesc & " " & xEnd & ".indd") -- no
					end if -- appellation = thisComponent
					
				end repeat --each indesignList searching for component type
				--end repeat --looking for rule_list and assigning
				--return pkgList
				--return ruleA
				set pkgListCt to count of pkgList
				--repeat with thisRule from 1 to count of rule_list
				--repeat with thisPDFpath from 1 to pkgListCt
				--	set quoted form of POSIX path of 
				--end repeat --set the order of pdfs
				--end repeat --each item in rule_list
				set destfile to thePath & jobNo & " " & clientID & " " & jobDesc & ".PDF"
			end repeat --each thisPart of rule_list
			(*		set cmd to "/usr/bin/joinPDF " & (quoted form of POSIX path of destfile) & PosixPDF
			do shell script cmd
*)
			
		end repeat --each clientRecord
		--return PosixPDF --> " '/Volumes/MrJr/zacrobs/1239 WYCH2 thisconstant OE.PDF' '/Volumes/MrJr/zacrobs/1239 WYCH2 thisconstant LR.PDF'" --works
	end tell --id4
end timeout
--end open
#################################################################################
-- #                                                                               #
-- #     SqliteClass.applescript                                                   #
-- #                                                                               #
-- #     author:   Craig Williams                                                  #
-- #     created:  2008-07-17                                                      #
-- #                                                                               #
-- #################################################################################
-- #                                                                               #
-- #     This program is free software: you can redistribute it and/or modify      #
-- #     it under the terms of the GNU General Public License as published by      #
-- #     the Free Software Foundation, either version 3 of the License, or         #
-- #     (at your option) any later version.                                       #
-- #                                                                               #
-- #     This program is distributed in the hope that it will be useful,           #
-- #     but WITHOUT ANY WARRANTY; without even the implied warranty of            #
-- #     MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the             #
-- #     GNU General Public License for more details.                              #
-- #                                                                               #
-- #     You should have received a copy of the GNU General Public License         #
-- #     along with this program.  If not, see <http://www.gnu.org/licenses/>;.     #
-- #                                                                               #
-- #################################################################################



(*
folder is created in ~/Library/Application Support/FOLDER_NAME
 
@property = FOLDER_NAME
@property = DATABASE_NAME
 
@function = create_db(table_name, column_names_array)
@function = sql_insert(table_name, the_values)
@function = sql_update(table_name, the_fields, the_values, search_field, search_value)
@function = sql_addColumn(table_name, col_name)
@function = sql_select(column_names_array, table_name, search_field, search_value)
@function = sql_select_all(column_names_array, table_name)
@function = sql_select_all_where(table_name, search_field, search_value)
@function = sql_delete(table_name, search_field, search_value)
@function = sql_delete_every_row(table_name)
@function = sql_delete_table(table_name)
*)



-- I create comma sep quoted string to enter into sqlite db
on returnCommaSepQuotedString(the_array)
	set return_string to ""
	if length of the_array > 1 then
		repeat with i from 1 to count of the_array
			set this_item to item i of the_array
			set return_string to return_string & "'" & this_item & "', "
		end repeat
		return text 1 thru -3 of return_string as string
	else
		return item 1 of the_array
	end if
end returnCommaSepQuotedString

on returnTabSepQuotedString(the_array)
	set return_string to ""
	set text item delimiters to tab
	set tab_array to text items of the_array
	--return length of tab_array
	if length of tab_array > 1 then
		--return length of the_array --tab delimited right --193
		repeat with i from 1 to count of tab_array
			set this_item to item i of tab_array
			set return_string to return_string & this_item & tab
		end repeat
		return text 1 thru -2 of return_string as string
	else
		return item 1 of the_array
	end if
end returnTabSepQuotedString

-- RETURN FILE PATH, HEAD, TAIL
on createFolderReturnFilePathHeadTail()
	set text item delimiters to ""
	set folder_name to "Databases:"
	set support_folder to (path to application support from user domain)
	tell application "Finder"
		set folder_path to (support_folder & folder_name) as string
		if not (exists folder folder_path) then
			make folder at support_folder with properties {name:folder_name}
		end if
	end tell
	set file_path to support_folder & "Databases" & ":" & database_name & ".db" as string
	set file_path to quoted form of POSIX path of file_path
	set loc to space & file_path & space
	set head to "sqlite3" & loc & quote
	set tail to quote
	return {file_path, head, tail}
end createFolderReturnFilePathHeadTail

-- CREATE DB tab
on create_dbtab(table_name, column_names_array)
	set column_names_string to my returnCommaSepQuotedString(column_names_array)
	set {file_path, head, tail} to my createFolderReturnFilePathHeadTail()
	set sql_statement to "create table if not exists " & table_name & "(" & column_names_string & "); "
	do shell script head & sql_statement & tail
end create_dbtab

-- CREATE DB orig
on create_db(table_name, column_names_array)
	set column_names_string to my returnTabSepQuotedString(column_names_array)
	set {file_path, head, tail} to my createFolderReturnFilePathHeadTail()
	set sql_statement to "create table if not exists " & table_name & "(" & column_names_string & "); "
	do shell script head & sql_statement & tail
end create_db


-- INSERT INTO DB sam
on sql_insert2(table_name, the_values)
	try
		set {file_path, head, tail} to my createFolderReturnFilePathHeadTail()
		set sql_statement to "insert into " & table_name & " values(" & the_values & "); "
		do shell script head & sql_statement & tail
	on error e
		display dialog "There was an error while inserting." & return & e
	end try
end sql_insert2

-- INSERT INTO DB orig
on sql_inserttab(table_name, the_values)
	try
		set {file_path, head, tail} to my createFolderReturnFilePathHeadTail()
		set the_values to my returnTabSepQuotedString(the_values)
		set sql_statement to "insert into " & table_name & " values(" & the_values & "); "
		do shell script head & sql_statement & tail
	on error e
		display dialog "There was an error while inserting." & return & e
	end try
end sql_inserttab

-- INSERT INTO DB orig
on sql_insert(table_name, the_values)
	try
		set {file_path, head, tail} to my createFolderReturnFilePathHeadTail()
		set the_values to my returnCommaSepQuotedString(the_values)
		set sql_statement to "insert into " & table_name & " values(" & the_values & "); "
		do shell script head & sql_statement & tail
	on error e
		display dialog "There was an error while inserting." & return & e
	end try
end sql_insert

-- UPDATE DB
on sql_update(table_name, the_fields, the_values, search_field, search_value)
	set {file_path, head, tail} to my createFolderReturnFilePathHeadTail()
	
	repeat with i from 1 to count of the_fields
		set this_item to item i of the_fields
		set sql_statement to ("UPDATE " & table_name & " set " & this_item & " = '" & Â
			item i of the_values & "' WHERE " & search_field & " = '" & search_value & "'; " as string)
		log sql_statement
		do shell script head & sql_statement & quote
	end repeat
end sql_update

-- ADD COLUMN
on sql_addColumn(table_name, col_name)
	set {file_path, head, tail} to my createFolderReturnFilePathHeadTail()
	set sql_statement to ("ALTER table " & table_name & " add " & col_name & "; " as string)
	do shell script head & sql_statement & quote
end sql_addColumn

-- SELECT
on sql_select(column_names_array, table_name, search_field, search_value)
	set {file_path, head, tail} to my createFolderReturnFilePathHeadTail()
	set column_names_string to my returnCommaSepQuotedString(column_names_array)
	set sql_statement to ("SELECT " & column_names_string & " FROM " & table_name & Â
		" WHERE " & search_field & " = '" & search_value & "'; " as string)
	return (do shell script head & sql_statement & quote)
end sql_select

-- SELECT ALL
on sql_select_all(table_name)
	set {file_path, head, tail} to my createFolderReturnFilePathHeadTail()
	set sql_statement to ("SELECT * FROM " & table_name & " ; " as string)
	set sql_execute to (do shell script head & sql_statement & quote)
	return my tidStuff(return, sql_execute)
end sql_select_all

-- SELECT ALL WHERE
on sql_select_all_where(table_name, search_field, search_value)
	set {file_path, head, tail} to my createFolderReturnFilePathHeadTail()
	set sql_statement to ("SELECT * FROM " & table_name & " WHERE " & Â
		search_field & " = " & search_value & " ; " as string)
	set sql_execute to (do shell script head & sql_statement & quote)
	return my tidStuff(return, sql_execute)
end sql_select_all_where

-- DELETE ONE ROW
on sql_delete(table_name, search_field, search_value)
	set {file_path, head, tail} to my createFolderReturnFilePathHeadTail()
	set sql_statement to ("DELETE FROM " & table_name & " WHERE " & Â
		search_field & " = " & search_value & " ; " as string)
	log sql_statement
	do shell script head & sql_statement & quote
end sql_delete

-- DELETE EVERY ROW
on sql_delete_every_row(table_name)
	set {file_path, head, tail} to my createFolderReturnFilePathHeadTail()
	set sql_statement to ("DELETE * FROM " & table_name & "; " as string)
	do shell script head & sql_statement & quote
end sql_delete_every_row

-- DELETE TABLE
on sql_delete_table(table_name)
	set {file_path, head, tail} to my createFolderReturnFilePathHeadTail()
	set sql_statement to ("DROP TABLE if exists " & table_name & "; " as string)
	do shell script head & sql_statement & quote
end sql_delete_table

-- TURN STRING INTO LIST
on tidStuff(paramHere, textHere)
	set OLDtid to AppleScript's text item delimiters
	set AppleScript's text item delimiters to paramHere
	set theItems to text items of textHere
	set AppleScript's text item delimiters to OLDtid
	return theItems
end tidStuff

---Nigel's csv converter
(* Assumes that the CSV text adheres to the convention:
	Records are delimited by LFs or CRLFs.
	The last record in the text may or may not be followed by an LF or CRLF.
	Fields in the same record are separated by commas.
	The last field in a record may or may not be followed by a comma.
	Trailing or leading spaces in unquoted fields are to be ignored.
	Fields enclosed in double quotes are to be taken verbatim,
		except for any included double-quote pairs, which are
		to be translated to double-quote characters.
		
	No variations are currently supported. Plain text is assumed throughout. *)

on csvToList(csvText)
	set LF to (ASCII character 10)
	
	script o -- Lists for fast access.
		property qdti : getTextItems(csvText, "\"")
		property currentRecord : {}
		property currentFields : missing value
		property recordList : {}
	end script
	
	-- o's qdti is a list of the CSV's text items, as delimited by double-quotes.
	-- Assuming the convention mentioned above, the number of items is always odd.
	-- Even-numbered items (if any) are quoted field values and don't need parsing.
	-- Odd-numbered items are everything else. Empty strings in odd-numbered slots
	-- (except at the beginning and end) indicate "quoted quotes" in quoted fields.
	
	set astid to AppleScript's text item delimiters
	set qdtiCount to (count o's qdti)
	set quoteInProgress to false
	considering case
		repeat with i from 1 to qdtiCount by 2 -- Parse odd-numbered items only.
			set thisBit to item i of o's qdti
			if ((count thisBit) > 0) or (i is qdtiCount) then
				-- This is either a non-empty string or the last item in the list, so it doesn't
				-- represent a quoted quote. Check if we've just been dealing with any.
				if (quoteInProgress) then
					-- All the parts of a quoted field containing quoted quotes have now been
					-- passed over. Coerce them together using a quote delimiter.
					set AppleScript's text item delimiters to "\""
					set thisField to (items a thru (i - 1) of o's qdti) as string
					-- Replace the reconstituted quoted quotes with literal quotes.
					set AppleScript's text item delimiters to "\"\""
					set thisField to thisField's text items
					set AppleScript's text item delimiters to "\""
					-- Store the field in the "current record" list and cancel the "quote in progress" flag.
					set end of o's currentRecord to thisField as string
					set quoteInProgress to false
				else if (i > 1) then
					-- The preceding, even-numbered item is a complete quoted field. Store it.
					set end of o's currentRecord to item (i - 1) of o's qdti
				end if
				
				-- Now get this item's comma-delimited text items, which are either non-quoted
				-- fields or stumps from the removal of quoted fields. Any that contain line
				-- feeds must be further split to end one record and start another. These could
				-- include multiple single-field records without comma separators.
				set o's currentFields to getTextItems(thisBit, ",")
				repeat with j from 1 to (count result)
					set thisField to item j of o's currentFields
					if (thisField contains LF) then
						-- This "field" contains one or more line endings. Split it into separate fields.
						set theseFields to thisField's paragraphs
						-- With each of these end-of-line fields except the last, complete the current
						-- record list and initialise another.
						set thisField to beginning of theseFields
						repeat with k from 2 to (count theseFields)
							set thisField to trim(thisField)
							if (thisField is not missing value) then set end of o's currentRecord to thisField
							set end of o's recordList to o's currentRecord
							set o's currentRecord to {}
							set thisField to item k of theseFields
						end repeat
					end if
					-- Store whatever's left of item j of o's currentFields in the current record list.
					set thisField to trim(thisField)
					if (thisField is not missing value) then set end of o's currentRecord to thisField
				end repeat
				
				-- Otherwise, this item IS an empty text representing a quoted quote.
			else if (quoteInProgress) then
				-- It's another quote in a field already identified as having one. Do nothing for now.
			else if (i > 1) then
				-- It's the first quoted quote in a quoted field. Note the index of the
				-- preceding even-numbered item (the first part of the field) and flag "quote in
				-- progress" so that the repeat idles past the remaining part(s) of the field.
				set a to i - 1
				set quoteInProgress to true
			end if
		end repeat
	end considering
	
	-- At the end of the repeat, store any remaining "current record".
	if (o's currentRecord is not {}) then set end of o's recordList to o's currentRecord
	set AppleScript's text item delimiters to astid
	
	return o's recordList
end csvToList

-- Get the possibly more than 4000 text items from a text.
on getTextItems(txt, delim)
	set astid to AppleScript's text item delimiters
	set AppleScript's text item delimiters to delim
	set tiCount to (count txt's text items)
	set textItems to {}
	repeat with i from 1 to tiCount by 4000
		set j to i + 3999
		if (j > tiCount) then set j to tiCount
		set textItems to textItems & text items i thru j of txt
	end repeat
	set AppleScript's text item delimiters to astid
	
	return textItems
end getTextItems

-- Trim any leading or trailing spaces from string.
on trim(txt)
	set len to (count txt)
	if (len > 0) then
		set b to 1
		set e to len
		repeat while (b < e) and (character b of txt is " ")
			set b to b + 1
		end repeat
		repeat while (e > b) and (character e of txt is " ")
			set e to e - 1
		end repeat
		if (b > 1) or (e < len) then set txt to text b thru e of txt
		if (txt is " ") then set txt to missing value
	else
		set txt to missing value
	end if
	
	return txt
end trim

------ ID handlers

on findChangePrefs(findText, changeText)
	tell application "Adobe InDesign CS5"
		set find text preferences to nothing
		set change text preferences to nothing
		set find what of find text preferences to findText
		set change to of change text preferences to changeText
	end tell
end findChangePrefs

on findGrepPrefs(findText, changeText)
	tell application "Adobe InDesign CS5"
		set find grep preferences to nothing
		set change grep preferences to nothing
		set find what of find grep preferences to findText
		set change to of change grep preferences to changeText
	end tell
end findGrepPrefs

on pdf81(aa, theDocument, theRange, thePreset, origRange, y0, x0, pageHgt, pageWid)
	set adate to short date string of (current date)
	tell application "Adobe InDesign CS5"
		set stripname to ""
		--set aa to name of document 1
		set ab to text -5 through end of aa
		if ab = ".indd" then
			set stripname to characters 1 thru -6 of aa as string
		else
			set stripname to aa as string
		end if
		--return stripname
		set theDocument to active document
		set theName to (stripname & ".PDF")
		set theRange to "1-" & (get count of pages of document 1)
		--return theRange
		--set thePath to file path of active document as text
		--return thePath
		set thePath to "MrJr:zacrobs:"
		set filePath to thePath & theName
		--set myFile to "Mac-PC Shared:*LA<>ATL:Grizzard Watermark low res.pdf"
		set myFile to "MrJr:@Art Staples:Grizzard, etc.:watermark:Grizzard Branding low res.pdf"
		set presetList to name of PDF export presets
		--set theChoice to choose from list presetList with prompt "Select PDF export preset" without multiple selections allowed and empty selection allowed
		--if theChoice is not false then
		set thePreset to "GrizProof"
		--	tell theDocument
		
		
		set origRange to page range of PDF export preferences
		set {y0, x0, pageHgt, pageWid} to bounds of page 1 of document 1
		(*
	--make border box
	repeat with i from 1 to (get count of pages of document 1)
		--tell page i of document 1
		make rectangle at page i of document 1 with properties {geometric bounds:{0.0, 0.0, pageHgt, pageWid}, label:"trashme", stroke weight:0.5, stroke alignment:inside alignment, fill color:"None", stroke transparency settings:"Normal 100%"}
		
		--end tell --page i
	end repeat
*)
		--return pageWid
		set whichStyle to ""
		--	try
		set myDocument to active document
		tell myDocument
			if exists (object reference of first color whose name is "**VARIABLE**") then
				set name of first color whose name is "**VARIABLE**" to "** VARIABLE LASER **"
			end if
			
			try
				set varObj to object reference of first color whose name contains "** VARIABLE LASER **"
			on error
				set varObj to make new color with properties {model:process, space:RGB, color value:{249.0, 0.0, 253.0}, name:"** VARIABLE LASER **"}
			end try
			
			if exists (object reference of first color whose name is "**LASER**") then
				set name of first color whose name is "**LASER**" to "** LASER **"
			end if
			try
				set lasObj to object reference of first color whose name contains "** LASER **"
			on error
				set lasObj to make new color with properties {model:process, space:RGB, color value:{0, 174, 239}, name:"** LASER **"}
			end try
			
			--return lasObj
			set origz to zero point
			set zero point to {0, 0}
			if not (exists layer "Border") then
				set myLayer to make layer with properties {name:"Border", ignore wrap:true, layer color:red, locked:false, visible:true, show guides:false}
				move layer "Border" to end of layer 1
			end if
			set borlock to ""
			set borlock to locked of layer "Border"
			set locked of layer "Border" to false
			--set theStyle to character style "counter"
			set rind to "#:"
			repeat with i from 1 to (get count of spreads)
				tell spread i
					--set counterreference to (a reference to (first text style range of every text frame whose applied character style is theStyle))
					--try
					set theGroups to every group
					repeat with aGroup in theGroups
						set TextFrames to every text frame of aGroup
						repeat with aFrame in TextFrames
							set TheText to text 1 of aFrame
							set rind to "#:"
							set rgb1 to "PRINT RGB KEY"
							if rind is in TheText then
								set whichStyle to "group"
								set hw to object reference of aFrame
								set hw3 to count of paragraphs of aFrame
								set hw2 to contents of paragraph hw3 of hw
								set colonPos to offset of ":" in hw2
								--return colonPos
								set hw4 to characters (colonPos + 2) thru -2 of hw2
								set prePost to characters 1 thru (colonPos + 1) of hw2
								set hw4 to hw4 as string
								--return prePost
								set counterreference to text from character (colonPos + 2) to (colonPos + 4) of hw2
								--return counterreference
								if counterreference contains " " then
									set counterreference to text from character (colonPos + 2) to (colonPos + 2) of hw2
								end if -- space
							end if -- rind in text
							if rgb1 is in TheText then
								set rgbRef to object reference of aFrame
								set rgbCograf to count of paragraphs of aFrame
								set rgbText to contents of paragraph rgbCograf of rgbRef
								--return rgbText
								set myStory to parent story of rgbRef
								tell every paragraph of myStory
									set fill color to lasObj
								end tell --every graf
							end if --rgb1
						end repeat -- frames
					end repeat --groups
					try --bail out of FPO box 
						set FPOpstoryObj to (parent story of first text frame whose label is "FPO color box")
						tell FPOpstoryObj
							set blueText to item 1 --of story 1
							--return text of blueText --this yields copy
							
							(*
					set text of blueText to "blink"
					set contents of insertion point 1 to "Tinker"
					end tell
					set contents of insertion point -1 to " Hall"
					tell every character of blueText
						set properties to {font:font "Impact"}
*)
							tell every paragraph
								set applied font to "Helvetica LT Std"
								set font style to "Regular"
								set contents to "DO NOT PRINT RGB KEYLINES OR TYPE, "
								set point size to 9
								set fill color to lasObj
							end tell --every graf
							tell insertion point -1
								set properties to {contents:"FPO color = Black LASER    ", fill color:lasObj}
								--works w/o font info
							end tell
							tell characters 36 through 62 of paragraph 1 -- contents of
								set applied font to "Helvetica LT Std"
								set font style to "Bold"
								--set font to "Helvetica LT Std Bold"
								set point size to 10
							end tell
							tell insertion point -1
								set properties to {contents:"FPO color = VARIABLE BLACK LASER", fill color:varObj}
							end tell
							--return count of characters of paragraph 1 --works -->1
							tell characters 64 through 94
								--set applied font to "Helvetica LT Std"
								set font style to "Bold"
								set point size to 10
								--set fill color to varObj
							end tell
							
						end tell --parent text frame
						--set tString to contents of counterreference -- as integer
						--return tString
						--on error
						--		display dialog "This requires the updated slit legend group." buttons {"OK"} default button 1
					end try --bail out of FPO box
					
					--New legend w/o groups
					if (first text frame whose label is "laser version") exists then
						set whichStyle to "broken1"
						set laserFrameObj to (first text frame whose label is "laser version") -- (parent of first text frame whose label is "laser version")
						--return laserFrameObj
						set theLaserTextObj to text 1 of laserFrameObj
						if rind is in theLaserTextObj then
							--beep 4
							set laserFrameParaCt to count of paragraphs of laserFrameObj
							--return laserFrameParaCt -->5
							repeat with thisGraf from 1 to laserFrameParaCt
								if rind is in paragraph thisGraf of laserFrameObj then
									
									--stuck here
									--	set colonLaseroffset to offset of ":" in contents of paragraph thisGraf of laserFrameObj				--tell paragraph hw3 of hw	
									set newGrafNo to thisGraf
									set Lw2 to contents of paragraph thisGraf of laserFrameObj
									set colonLaseroffset to offset of ":" in Lw2
									--return colonLaseroffset -->18
									set Lw4 to characters (colonLaseroffset + 2) thru -2 of Lw2
									set Lw4 to Lw4 as string
									set counterreference to text from character (colonLaseroffset + 2) to (colonLaseroffset + 4) of Lw2
									if counterreference contains " " then
										set counterreference to text from character (colonLaseroffset + 2) to (colonLaseroffset + 2) of Lw2
										set prePost to characters 1 thru (colonLaseroffset + 1) of Lw2
									end if --blank space in number
								end if --rind find
							end repeat --each graf in Laser box
							
							--set colonLaseroffset to offset of ":" in theLaserTextObj
							--set colonLaseroffsetGraf to parent of character colonLaseroffset of theLaserTextObj
						end if
					end if -- broken legend exists
					set allPrepFrame to (every text frame whose label is "preprint version")
					--return (count of paragraphs of text 1 of item 1 of aaa)
					if (exists (first text frame whose label is "preprint version")) and (whichStyle ­ "grouped") and (whichStyle ­ "broken1") and ((count of paragraphs of text 1 of item 1 of allPrepFrame) > 1) then
						set whichStyle to "print"
						set printFrameObj to (first text frame whose label is "preprint version") -- (parent of first text frame whose label is "preprint version")
						--return printFrameObj
						set thePrintTextObj to text 1 of printFrameObj
						if rind is in thePrintTextObj then
							--beep 4
							set printFrameParaCt to count of paragraphs of printFrameObj
							--return printFrameParaCt -->5
							repeat with thisGraf from 1 to printFrameParaCt
								if rind is in paragraph thisGraf of printFrameObj then
									
									--stuck here
									--	set colonLaseroffset to offset of ":" in contents of paragraph thisGraf of laserFrameObj				--tell paragraph hw3 of hw	
									set newGrafNo to thisGraf
									set Pw2 to contents of paragraph thisGraf of printFrameObj
									set colonPrintoffset to offset of ":" in Pw2
									--return colonPrintoffset -->18
									set Pw4 to characters (colonPrintoffset + 2) thru -2 of Pw2
									set Pw4 to Pw4 as string
									set counterreference to text from character (colonPrintoffset + 2) to (colonPrintoffset + 4) of Pw2
									if counterreference contains " " then
										set counterreference to text from character (colonPrintoffset + 2) to (colonPrintoffset + 2) of Pw2
										set prePost to characters 1 thru (colonPrintoffset + 1) of Pw2
									end if --blank space in number
								end if --rind find
							end repeat --each graf in Laser box
						end if --rindfind
					end if --exists print
					(*
set laserStoryObj to (parent story of first text frame whose label is "laser version")
				tell laserStoryObj
					set laserLotText to item 1
					--set text of laserLotText to "Harrow"
					set colonLaseroffset to offset of ":" in text of laserLotText
					end tell --laser box parent story
				--	return text of laserLotText
				return colonLaseroffset
*)
					set howie to counterreference
					try
						set howie to counterreference as integer
					end try --integer
					--end try
					set myRectangle to make rectangle with properties {geometric bounds:{-1.0, 0.0, -0.2495, 3}, stroke weight:0, stroke color:swatch "None" of myDocument, item layer:layer 1 of myDocument, label:"bixby"}
					tell myRectangle
						set myGraphic to place (myFile as string)
						set myGraphic to item 1 of myGraphic
						fit given proportionally
						fit given center content
					end tell --myRect
					--end try
					
					--return howie
					--return theSpread
					--	end tell --pg i
					set counterNum to text returned of (display dialog "Round # for page " & i & "?" buttons {"OK"} default answer (howie) default button "OK" with icon 1 giving up after 6)
					
					if counterNum is not equal to (howie as string) then
						--set contents of counterreference to (counterNum & tab & adate)
						if whichStyle ­ "" then
							if whichStyle is "group" then
								tell paragraph hw3 of hw
									set contents to (prePost & "" & counterNum & "  Date: " & adate as string)
									--of (text characters (colonPos + 2) thru -2)
								end tell --hw
							else if whichStyle is "broken1" then
								tell paragraph newGrafNo of laserFrameObj
									set contents to ((prePost & "" & counterNum & "  Date: " & adate as string) & return)
								end tell -- new laser box
								-- end if -- grouped or broken
							else
								tell paragraph newGrafNo of printFrameObj
									set contents to ((prePost & "" & counterNum & "  Date: " & adate as string) & return)
								end tell -- new print box
							end if --
							--set contents of counterreference to (tString as text)
						end if --whichStyle not empty
					end if
					--return counterreference
					
					(*
on error
					display dialog "perhaps page " & i & " doesn't have a legend "
					
				end try
*)
				end tell --spread i
			end repeat
			set zero point to origz
		end tell --doc1
		
		
		set properties of PDF export preferences to {page range:theRange}
		export theDocument format PDF type to filePath using PDF export preset thePreset
		--set page range of PDF export preferences to origRange
		try
			delete (every rectangle of document 1 whose label = "trashme")
		end try
		try
			delete (every rectangle of document 1 whose label contains "bixby")
		end try
		
		set locked of layer "Border" of document 1 to borlock
		
	end tell --id4
end pdf81
--*************************

--POSIX path from Bruce Phillips. Needs a POSIX path
on escapePOSIX(somePath)
	set astid to AppleScript's text item delimiters
	set escapeNeeded to {"\\", space, ASCII character 34, tab, "(", ")", "[", "]", "{", "}", "<", ">", "|", "!", "?", "$", "~", "^", "*", "`", "'", ";", ASCII character 27}
	
	try
		repeat with thisChar in escapeNeeded
			set AppleScript's text item delimiters to {thisChar}
			set somePath to text items of somePath
			
			set AppleScript's text item delimiters to {"\\" & thisChar}
			set somePath to "" & somePath
		end repeat
	end try
	
	set AppleScript's text item delimiters to astid
	return somePath
end escapePOSIX
