=begin
From BIFF8 on, strings are always stored using UTF-16LE  text encoding. The
character  array  is  a  sequence  of  16-bit  values4.  Additionally it is
possible  to  use  a  compressed  format, which omits the high bytes of all
characters, if they are all zero.

The following tables describe the standard format of the entire string, but
in many records the strings differ from this format. This will be mentioned
separately. It is possible (but not required) to store Rich-Text formatting
information  and  Asian  phonetic information inside a Unicode string. This
results  in  four  different  ways  to  store a string. The character array
is not zero-terminated.

The  string  consists  of  the  character count (as usual an 8-bit value or
a  16-bit value), option flags, the character array and optional formatting
information.  If the string is empty, sometimes the option flags field will
not occur. This is mentioned at the respective place.

Offset  Size    Contents
0       1 or 2  Length of the string (character count, ln)
1 or 2  1       Option flags:
                  Bit   Mask Contents
                  0     01H  Character compression (ccompr):
                               0 = Compressed (8-bit characters)
                               1 = Uncompressed (16-bit characters)
                  2     04H  Asian phonetic settings (phonetic):
                               0 = Does not contain Asian phonetic settings
                               1 = Contains Asian phonetic settings
                  3     08H  Rich-Text settings (richtext):
                               0 = Does not contain Rich-Text settings
                               1 = Contains Rich-Text settings
[2 or 3] 2      (optional, only if richtext=1) Number of Rich-Text formatting runs (rt)
[var.]   4      (optional, only if phonetic=1) Size of Asian phonetic settings block (in bytes, sz)
var.     ln or 
         2·ln   Character array (8-bit characters or 16-bit characters, dependent on ccompr)
[var.]   4·rt   (optional, only if richtext=1) List of rt formatting runs 
[var.]   sz     (optional, only if phonetic=1) Asian Phonetic Settings Block 
=end


require 'iconv'
module Excel
    module UnicodeUtils
	
	def u2ints(str)
	    Excel::ICONV[:to_unicode].iconv(str).unpack('S*')
	end
	
	def u2bytes(str)
	    Excel::ICONV[:to_unicode].iconv(str)
	end
	
	def upack2(str)
	    begin
		ustr = Excel::ICONV[:check_ascii].iconv(str)
		[str.length, 0].pack('vC')+str
	    rescue Iconv::IllegalSequence
		ustr = u2bytes(str)
		[ustr.length / 2, 1].pack('vC')+ustr
	    end
	end
	
	def upack1(str)
	    begin
		ustr = Excel::ICONV[:check_ascii].iconv(str)
		[str.length, 0].pack('CC')+str
	    rescue Iconv::IllegalSequence
		ustr = u2bytes(str)
		[ustr.length / 2, 1].pack('CC')+ustr
	    end
	end
    end
end
