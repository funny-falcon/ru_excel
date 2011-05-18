dir = File.dirname(__FILE__)
require dir+'/biff_records'
require dir+'/deco'

module Excel
  class Column
    def initialize(indx, parent_sheet)
        @index = indx
        @parent = parent_sheet
        @parent_wb = parent_sheet.parent
        @xf_index = 0x0F
        
        @width = 0x0B92
        @hidden = 0
        @level = 0
        @collapse = 0
    end
    attr_accessor :level, :hidden, :width, :collapse


    def get_biff_record
        options =  (@hidden & 0x01) << 0
        options |= (@level & 0x07) << 8
        options |= (@collapse & 0x01) << 12
        
        BiffRecord.colInfoRecord(@index, @index, @width, @xf_index, options)
    end
  end
end        
        
        
