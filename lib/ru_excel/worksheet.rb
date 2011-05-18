dir = File.dirname(__FILE__)
require dir+'/biff_records'
require dir+'/bitmap'
require dir+'/formatting'
require dir+'/style'
require dir+'/deco'
require dir+'/row'
require dir+'/column'
require dir+'/cell'
require dir+'/workbook'


=begin
            BOF
            UNCALCED
            INDEX
            Calculation Settings Block
            PRINTHEADERS
            PRINTGRIDLINES
            GRIDSET
            GUTS
            DEFAULTROWHEIGHT
            WSBOOL
            Page Settings Block
            Worksheet Protection Block
            DEFCOLWIDTH
            COLINFO
            SORT
            DIMENSIONS
            Row Blocks
            WINDOW2
            SCL
            PANE
            SELECTION
            STANDARDWIDTH
            MERGEDCELLS
            LABELRANGES
            PHONETIC
            Conditional Formatting Table
            Hyperlink Table
            Data Validity Table
            SHEETLAYOUT (BIFF8X only)
            SHEETPROTECTION (BIFF8X only)
            RANGEPROTECTION (BIFF8X only)
            EOF
=end

module Excel
  class Worksheet
    extend Deco
    #################################################################
    ## Constructor
    #################################################################
    def initialize(sheetname, parent_book)
        @name = sheetname
        @parent = parent_book

        @rows = {}
        @cols = {}
        @merged_ranges = []
        @bmp_rec = ''

        @show_formulas = 0
        @show_grid = 1
        @show_headers = 1
        @panes_frozen = 0
        @show_empty_as_zero = 1
        @auto_colour_grid = 1
        @cols_right_to_left = 0
        @show_outline = 1
        @remove_splits = 0
        @selected = 0
        @hidden = 0
        @page_preview = 0

        @first_visible_row = 0
        @first_visible_col = 0
        @grid_colour = 0x40
        @preview_magn = 0
        @normal_magn = 0

        @vert_split_pos = nil
        @horz_split_pos = nil
        @vert_split_first_visible = nil
        @horz_split_first_visible = nil
        @split_active_pane = 

        @row_gut_width = 0
        @col_gut_height = 0

        @show_auto_page_breaks = 1
        @dialogue_sheet = 0
        @auto_style_outline = 0
        @outline_below = 0
        @outline_right = 0
        @fit_num_pages = 0
        @show_row_outline = 1
        @show_col_outline = 1
        @alt_expr_eval = 0
        @alt_formula_entries = 0

        @row_default_height = 0x00FF
        @col_default_width = 0x0008

        @calc_mode = 1
        @calc_count = 0x0064
        @RC_ref_mode = 1
        @iterations_on = 0
        @delta = 0.001
        @save_recalc = 0

        @print_headers = 0
        @print_grid = 0
        @grid_set = 1
        @vert_page_breaks = []
        @horz_page_breaks = []
        @header_str = '&P'
        @footer_str = '&F'
        @print_centered_vert = 0
        @print_centered_horz = 1
        @left_margin = 0.3 #0.5
        @right_margin = 0.3 #0.5
        @top_margin = 0.61 #1.0
        @bottom_margin = 0.37 #1.0
        @paper_size_code = 9 # A4
        @print_scaling = 100
        @start_page_number = 1
        @fit_width_to_pages = 1
        @fit_height_to_pages = 1
        @print_in_rows = 1
        @portrait = 1
        @print_not_colour = 0
        @print_draft = 0
        @print_notes = 0
        @print_notes_at_end = 0
        @print_omit_errors = 0
        @print_hres = 0x012C # 300 dpi
        @print_vres = 0x012C # 300 dpi
        @header_margin = 0.1
        @footer_margin = 0.1
        @copies_num = 1

        @wnd_protect = 0
        @obj_protect = 0
        @protect = 0
        @scen_protect = 0
        @password = ''
    end
    #################################################################
    ## Properties, "getters", "setters"
    #################################################################
    def calc_mode=(value)
        @calc_mode = value & 0x03
    end
    def calc_mode
        return @calc_mode
    end
    string_accessor :name, :header_str, :footer_str, :password
    float_accessor  :delta, :left_margin, :right_margin, :top_margin, :bottom_margin
    float_accessor  :header_margin, :footer_margin
    array_accessor  :vert_page_breaks, :horz_page_breaks

    attr_reader :parent, :rows, :cols, :merged_ranges, :bmp_rec
    bool_int_accessor :print_colour, :show_formulas, :show_grid, :show_headers
    bool_int_accessor :panes_frozen, :show_empty_as_zero, :auto_colour_grid
    bool_int_accessor :cols_right_to_left, :show_outline, :remove_splits, :selected
    bool_int_accessor :hidden, :page_preview, :show_auto_page_breaks
    bool_int_accessor :dialogue_sheet, :auto_style_outline, :outline_below
    bool_int_accessor :outline_right, :show_row_outline, :show_col_outline
    bool_int_accessor :alt_expr_eval, :alt_formula_entries, :RC_ref_mode
    bool_int_accessor :iterations_on, :save_recalc, :print_headers, :print_grid
    bool_int_accessor :print_centered_vert, :print_centered_horz, :print_in_rows
    bool_int_accessor :portrait, :print_not_colour, :print_draft, :print_notes
    bool_int_accessor :print_notes_at_end, :print_omit_errors, :wnd_protect
    bool_int_accessor :obj_protect, :protect, :scen_protect

    int_accessor :first_visible_row, :first_visible_col, :grid_colour, :preview_magn
    int_accessor :normal_magn, :fit_num_pages, :row_default_height, :print_scaling
    int_accessor :col_default_width, :calc_count, :paper_size_code, :print_hres
    int_accessor :start_page_number, :fit_width_to_pages, :fit_height_to_pages
    int_accessor :print_vres, :copies_num
    
    absint_accessor :vert_split_pos, :horz_split_pos, :vert_split_first_visible 
    absint_accessor :horz_split_first_visible

    attr_reader :parent

    ##################################################################
    ## Methods
    ##################################################################

    def write(r, c, label="", style=XFStyle.new())
        row(r).write(c, label, style)
    end

    def merge(r1, r2, c1, c2, style=Style.XFStyle())
        for r in r1..r2
            row(r).write_blanks(c1, c2,  style)
        end
        @merged_ranges << [r1, r2, c1, c2]
    end

    def write_merge(r1, r2, c1, c2, label="", style=XFStyle.new())
        merge(r1, r2, c1, c2, style)
        write(r1, c1,  label, style)
    end

    def insert_bitmap(filename, row, col, x = 0, y = 0, scale_x = 1, scale_y = 1)
        bmp = BiffRecord.imDataBmpRecord(filename)
        obj = BiffRecord.objBmpRecord(row, col, self, bmp, x, y, scale_x, scale_y)

        @bmp_rec += obj + bmp
    end
    def col(indx)
        @cols[indx] ||= Column.new(indx, self)
    end

    def row(indx)
        @rows[indx] ||= Row.new(indx, self)
    end

    def row_height(row) # in pixels
        if (row = @rows[row])
            return row.get_height_in_pixels()
        else
            return 17
        end
    end

    def col_width(col) # in pixels
        #if col in @cols
        #    return @cols[col].width_in_pixels()
        #else
            return 64
    end

    def labels_count
        result = 0
        for r, row in @rows
            result += row.get_str_count()
        end
        return result
    end

    ##################################################################
    ## BIFF records generation
    ##################################################################

    def _bof_rec
        BiffRecord.biff8BOFRecord(BiffRecord::WORKSHEET)
    end

    def _guts_rec
        row_visible_levels = 0
        if @rows.length != 0
            row_visible_levels = @rows.map{|k,r| r.level}.max + 1
        end

        col_visible_levels = 0
        if @cols.length != 0
            col_visible_levels = @cols.map{|k,c| c.level}.max + 1
        end

        BiffRecord.gutsRecord(@row_gut_width, @col_gut_height, row_visible_levels, col_visible_levels)
    end

    def _wsbool_rec
        options = 0x00
        options |= (@show_auto_page_breaks & 0x01) << 0
        options |= (@dialogue_sheet & 0x01) << 4
        options |= (@auto_style_outline & 0x01) << 5
        options |= (@outline_below & 0x01) << 6
        options |= (@outline_right & 0x01) << 7
        options |= (@fit_num_pages & 0x01) << 8
        options |= (@show_row_outline & 0x01) << 10
        options |= (@show_col_outline & 0x01) << 11
        options |= (@alt_expr_eval & 0x01) << 14
        options |= (@alt_formula_entries & 0x01) << 15

        BiffRecord.wsBoolRecord(options)
    end

    def _eof_rec
        BiffRecord.eofRecord()
    end

    def _colinfo_rec
        @cols.map{|k, col| col.get_biff_record}.join('')
    end

    def _dimensions_rec
        first_used_row = 0
        last_used_row = 0
        first_used_col = 0
        last_used_col = 0
        if @rows.length > 0
            first_used_row = @rows.min[0]
            last_used_row = @rows.max[0]
            first_used_col = 0xFFFFFFFF
            last_used_col = 0
            for k,r in @rows
                _min = r.get_min_col()
                _max = r.get_max_col()
                if _min < first_used_col
                    first_used_col = _min
                end
                if _max > last_used_col
                    last_used_col = _max
                end
            end
        end

        BiffRecord.dimensionsRecord(first_used_row, last_used_row, first_used_col, last_used_col)
    end

    def _window2_rec
        options = 0
        options |= (@show_formulas        & 0x01) << 0
        options |= (@show_grid            & 0x01) << 1
        options |= (@show_headers         & 0x01) << 2
        options |= (@panes_frozen         & 0x01) << 3
        options |= (@show_empty_as_zero   & 0x01) << 4
        options |= (@auto_colour_grid     & 0x01) << 5
        options |= (@cols_right_to_left   & 0x01) << 6
        options |= (@show_outline         & 0x01) << 7
        options |= (@remove_splits        & 0x01) << 8
        options |= (@selected             & 0x01) << 9
        options |= (@hidden               & 0x01) << 10
        options |= (@page_preview         & 0x01) << 11

        BiffRecord.window2Record(options, @first_visible_row, @first_visible_col,
                                        @grid_colour,
                                        @preview_magn, @normal_magn)
    end

    def _panes_rec
        return "" if @vert_split_pos == nil && @horz_split_pos == nil

        @vert_split_pos ||= 0
        @horz_split_pos ||= 0

        if @panes_frozen
            @vert_split_first_visible ||= @vert_split_pos
            @horz_split_first_visible ||= @horz_split_pos
        else
            @vert_split_first_visible ||= 0
            @horz_split_first_visible ||= 0
            # inspired by pyXLWriter
            @horz_split_pos = 20*@horz_split_pos + 255
            @vert_split_pos = 113.879*@vert_split_pos + 390
        end

        @split_active_pane = (@vert_split_pos > 0 ? 0 : 2) + (@horz_split_pos > 0 ? 0 : 1)

        BiffRecord.panesRecord(@vert_split_pos,
                                         @horz_split_pos,
                                         @horz_split_first_visible,
                                         @vert_split_first_visible,
                                         @split_active_pane)
    end

    def _row_blocks_rec
        # this function takes almost 99% of overall execution time 
        # when file is saved
        # return '' 
        result = []
        for k, r in @rows
            result << r.get_row_biff_data
            result << r.get_cells_biff_data
        end
        result.join('')
    end

    def _merged_rec
        BiffRecord.mergedCellsRecord(@merged_ranges)
    end

    def _bitmaps_rec
        @bmp_rec
    end

    def _calc_settings_rec
        result = ''
        result << BiffRecord.calcModeRecord(@calc_mode & 0x01)
        result << BiffRecord.calcCountRecord(@calc_count & 0xFFFF)
        result << BiffRecord.refModeRecord(@RC_ref_mode & 0x01)
        result << BiffRecord.iterationRecord(@iterations_on & 0x01)
        result << BiffRecord.deltaRecord(@delta)
        result << BiffRecord.saveRecalcRecord(@save_recalc & 0x01)
        result
    end

    def _print_settings_rec
        result = ''
        result << BiffRecord.printHeadersRecord(@print_headers)
        result << BiffRecord.printGridLinesRecord(@print_grid)
        result << BiffRecord.gridSetRecord(@grid_set)
        result << BiffRecord.horizontalPageBreaksRecord(@horz_page_breaks)
        result << BiffRecord.verticalPageBreaksRecord(@vert_page_breaks)
        result << BiffRecord.headerRecord(@header_str)
        result << BiffRecord.footerRecord(@footer_str)
        result << BiffRecord.hCenterRecord(@print_centered_horz)
        result << BiffRecord.vCenterRecord(@print_centered_vert)
        result << BiffRecord.leftMarginRecord(@left_margin)
        result << BiffRecord.rightMarginRecord(@right_margin)
        result << BiffRecord.topMarginRecord(@top_margin)
        result << BiffRecord.bottomMarginRecord(@bottom_margin)

        setup_page_options =  (@print_in_rows & 0x01) << 0
        setup_page_options |=  (@portrait & 0x01) << 1
        setup_page_options |=  (0x00 & 0x01) << 2
        setup_page_options |=  (@print_not_colour & 0x01) << 3
        setup_page_options |=  (@print_draft & 0x01) << 4
        setup_page_options |=  (@print_notes & 0x01) << 5
        setup_page_options |=  (0x00 & 0x01) << 6
        setup_page_options |=  (0x01 & 0x01) << 7
        setup_page_options |=  (@print_notes_at_end & 0x01) << 9
        setup_page_options |=  (@print_omit_errors & 0x03) << 10

        result << BiffRecord.setupPageRecord(@paper_size_code,
                                @print_scaling,
                                @start_page_number,
                                @fit_width_to_pages,
                                @fit_height_to_pages,
                                setup_page_options,
                                @print_hres,
                                @print_vres,
                                @header_margin,
                                @footer_margin,
                                @copies_num)
        result
    end

    def _protection_rec
        result = ''
        result << BiffRecord.protectRecord(@protect)
        result << BiffRecord.scenProtectRecord(@scen_protect)
        result << BiffRecord.windowProtectRecord(@wnd_protect)
        result << BiffRecord.objectProtectRecord(@obj_protect)
        result << BiffRecord.passwordRecord(@password)
        result
    end

    def get_biff_data
        result = ''
        result << _bof_rec()
        result << _calc_settings_rec()
        result << _guts_rec()
        result << _wsbool_rec()
        result << _colinfo_rec()
        result << _dimensions_rec()
        result << _print_settings_rec()
        result << _protection_rec()
        result << _row_blocks_rec()
        result << _merged_rec()
        result << _bitmaps_rec()
        result << _window2_rec()
        result << _panes_rec()
        result << _eof_rec()
        result
    end
  end
end



