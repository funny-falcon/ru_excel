dir = File.dirname(__FILE__)
require dir+'/biff_records'
require dir+'/worksheet'
require dir+'/style'
require 'date'

module Excel
  class Row

    #################################################################
    ## Constructor
    #################################################################
    def initialize(index, parent_sheet)
        @idx = index
        @parent = parent_sheet
        @parent_wb = parent_sheet.parent
        @cells = []
        @min_col_idx = 0
        @max_col_idx = 0
        @total_str = 0
        @xf_index = 0x0F
        @has_default_format = 0
        @height_in_pixels = 0x11
        
        @height = 0x00FF
        @has_default_height = 0x00
        @level = 0
        @collapse = 0
        @hidden = 0
        @space_above = 0
        @space_below = 0
    end
    
    attr_accessor :level

    def _adjust_height(style)
        twips = style.font.height
        points = twips.to_f/20.0
        # Cell height in pixels can be calcuted by following approx. formula
        # cell height in pixels = font height in points * 83/50 + 2/5
        # It works when screen resolution is 96 dpi 
        pix = (points*83.0/50.0 + 2.0/5.0.to_i).round
        if pix > @height_in_pixels
            @height_in_pixels = pix
        end
    end


    def _adjust_bound_col_idx(*args)
        for arg in args
            if arg < @min_col_idx
                @min_col_idx = arg
            elsif arg > @max_col_idx
                @max_col_idx = arg
            end
        end
    end

    EPOCH = Date.new(1899, 12, 31)
    TIME_EPOCH = Time.mktime(1902, 1, 1)
    TIME_EPOCH_ADD = (Date.new(1902, 1, 1) - Date.new(1899, 12, 31))
    def _excel_date_dt(date)
        if Date === date
            xldate = (date - EPOCH).to_f + date.offset.to_f
        elsif Time === date
            xldate = (date - TIME_EPOCH + 
                      date.utc_offset - TIME_EPOCH.utc_offset).to_f / (24 * 60 * 60) +
                     TIME_EPOCH_ADD
        else
            raise date.inspect+' is not a date'
        end
        # Add a day for Excel's missing leap day in 1900
        xldate += 1 if xldate > 59
        xldate
    end

    def get_height_in_pixels
        @height_in_pixels
    end


    def set_style(style)
        _adjust_height(style)
        @xf_index = @parent_wb.add_style(style)
    end

            
    def get_xf_index
        @xf_index
    end

    
    def get_cells_count
        @cells.length
    end

    
    def get_min_col
        @min_col_idx
    end

        
    def get_max_col
        @min_col_idx
    end

        
    def get_str_count
        @total_str
    end


    def get_row_biff_data
        height_options = (@height & 0x07FFF) 
        height_options |= (@has_default_height & 0x01) << 15

        options =  (@level & 0x07) << 0
        options |= (@collapse & 0x01) << 4
        options |= (@hidden & 0x01) << 5
        options |= (0x00 & 0x01) << 6
        options |= (0x01 & 0x01) << 8
        if @xf_index != 0x0F
            options |= (0x01 & 0x01) << 7
        else
            options |= (0x00 & 0x01) << 7
        end
        options |= (@xf_index & 0x0FFF) << 16 
        options |= (0x00 & @space_above) << 28
        options |= (0x00 & @space_below) << 29
        
        BiffRecord.rowRecord(@idx, @min_col_idx, @max_col_idx, height_options, options)
    end

    def get_cells_biff_data
        #@cells.map{|c| c.get_biff_data()}.join('')
        @cells.join('')
    end


    def get_index
        @idx
    end


    def write(col, label, style)
        _adjust_height(style)
        _adjust_bound_col_idx(col)
        if String === label
            if label.length > 0
                @cells << Cell.strCell(self, col, @parent_wb.add_style(style), @parent_wb.add_str(label))
                @total_str += 1
            else
                @cells << Cell.blankCell(self, col, @parent_wb.add_style(style))
            end
        elsif Numeric === label
            @cells << Cell.numberCell(self, col, @parent_wb.add_style(style), label)
        elsif Date === label or Time === label
            @cells << Cell.numberCell(self, col, @parent_wb.add_style(style), _excel_date_dt(label))
        else
            @cells << Cell.formulaCell(self, col, @parent_wb.add_style(style), label)
        end
    end

    def write_blanks(c1, c2, style)
        _adjust_height(style)
        _adjust_bound_col_idx(c1, c2)
        @cells << Cell.mulBlankCell(self, c1, c2, @parent_wb.add_style(style))
    end
  end
end