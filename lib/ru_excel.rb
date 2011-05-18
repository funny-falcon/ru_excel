require 'iconv'
module Excel
    ICONV = {}
    def self.encoding
	@encoding
    end
    def self.encoding=(enc)
	@encoding = enc
	(ICONV[:to_unicode] = Iconv.new('utf16le', enc)).iconv('z')
	(ICONV[:from_unicode] = Iconv.new(enc, 'utf16le')).iconv("z\0")
    end
    self.encoding= 'cp1251'    
    (ICONV[:check_ascii] = Iconv.new('ascii','ascii')).iconv('z')
    (ICONV[:check_unicode] = Iconv.new('utf16le','utf16le')).iconv("z\0")
end
require 'ru_excel/workbook'
