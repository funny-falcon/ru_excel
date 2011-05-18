=begin
Record Order in BIFF8
  Workbook Globals Substream
      BOF Type = workbook globals
      Interface Header
      MMS
      Interface End
      WRITEACCESS
      CODEPAGE
      DSF
      TABID
      FNGROUPCOUNT
      Workbook Protection Block
            WINDOWPROTECT
            PROTECT
            PASSWORD
            PROT4REV
            PROT4REVPASS
      BACKUP
      HIDEOBJ 
      WINDOW1 
      DATEMODE 
      PRECISION
      REFRESHALL
      BOOKBOOL 
      FONT +
      FORMAT *
      XF +
      STYLE +
    ? PALETTE
      USESELFS
    
      BOUNDSHEET +
    
      COUNTRY 
    ? Link Table 
      SST 
      ExtSST
      EOF
=end
dir = File.dirname(__FILE__)
require dir + '/style'
require dir + '/deco'
require dir + '/worksheet'
require dir + '/compound_doc'

module Excel
  class Workbook
    extend Deco
    #################################################################
    ## Constructor
    #################################################################
    def initialize
        @owner = 'None'
        @country_code = 0x07 
        @wnd_protect = 0
        @obj_protect = 0
        @protect = 0
        @backup_on_save = 0
        # for WINDOW1 record
        @hpos_twips = 0x01E0
        @vpos_twips = 0x005A
        @width_twips = 0x3FCF
        @height_twips = 0x2A4E
        
        @active_sheet = 0
        @first_tab_index = 0
        @selected_tabs = 0x01
        @tab_width_twips = 0x0258
        
        @wnd_hidden = 0
        @wnd_mini = 0
        @hscroll_visible = 1
        @vscroll_visible = 1
        @tabs_visible = 1

        @styles = StyleCollection.new()
         
        @dates_1904 = 0
        @use_cell_values = 1
        
        @sst = SharedStringTable.new()
        
        @worksheets = []
    end
    #################################################################
    ## Properties, "getters", "setters"
    #################################################################

    string_accessor :owner
    int_accessor :country_code
    bool_accessor :wnd_protect, :protect, :backup_on_save

    short_accessor :hpos, :vpos, :width, :height, :tab_width

    def active_sheet=(value)
        @first_tab_index = @active_sheet = value.to_i & 0xFFFF
    end
    attr_reader :active_sheet
    
    def wnd_visible=(value)
        @wnd_hidden = !value
    end
    def wnd_visible
        !@wnd_hidden
    end
    bool_accessor :wnd_visible

    bool_int_accessor :wnd_mini, :hscroll_visible, :vscroll_visible, :obj_protect
    bool_int_accessor :tabs_visible, :dates_1904, :use_cell_values
    
    def default_style
        @styles.default_style
    end

    ##################################################################
    ## Methods
    ##################################################################

    def add_style(style)
       @styles.add(style)
    end

    def add_str(s)
       @sst.add_str(s)
    end

    def str_index(s)
        @sst.str_index(s)
    end

    def add_sheet(sheetname)
        @worksheets << Worksheet.new(sheetname, self)
        @worksheets[-1]
    end

    def get_sheet(sheetnum)
        @worksheets[sheetnum]
    end
    
    ##################################################################
    ## BIFF records generation
    ##################################################################

    def _bof_rec
        BiffRecord.biff8BOFRecord(BiffRecord::BOOK_GLOBAL)
    end

    def _eof_rec
        BiffRecord.eofRecord()
    end
        
    def _intf_hdr_rec
        BiffRecord.interaceHdrRecord()
    end

    def _intf_end_rec
        BiffRecord.interaceEndRecord()
    end

    def _intf_mms_rec
        BiffRecord.mmsRecord()
    end

    def _write_access_rec
        BiffRecord.writeAccessRecord(@owner)
    end

    def _wnd_protect_rec
        BiffRecord.windowProtectRecord(@wnd_protect)
    end

    def _obj_protect_rec
        BiffRecord.objectProtectRecord(@obj_protect)
    end

    def _protect_rec
        BiffRecord.protectRecord(@protect)
    end

    def _password_rec
        BiffRecord.passwordRecord()
    end

    def _prot4rev_rec
        BiffRecord.prot4RevRecord()
    end

    def _prot4rev_pass_rec
        BiffRecord.prot4RevPassRecord()
    end

    def _backup_rec
        BiffRecord.backupRecord(@backup_on_save)
    end
        
    def _hide_obj_rec
        BiffRecord.hideObjRecord()
    end
        
    def _window1_rec
        flags = 0
        flags |= (@wnd_hidden) << 0
        flags |= (@wnd_mini) << 1
        flags |= (@hscroll_visible) << 3
        flags |= (@vscroll_visible) << 4
        flags |= (@tabs_visible) << 5
        
        BiffRecord.window1Record(@hpos_twips, @vpos_twips, 
                                @width_twips, @height_twips, 
                                flags,
                                @active_sheet, @first_tab_index, 
                                @selected_tabs, @tab_width_twips)
    end
    
    def _codepage_rec
        BiffRecord.codepageBiff8Record()
    end
        
    def _country_rec
        BiffRecord.countryRecord(@country_code, @country_code)
    end
     
    def _dsf_rec
        BiffRecord.dsfRecord()
    end
    
    def _tabid_rec
        BiffRecord.tabIDRecord(@worksheets.length)
    end
    def _fngroupcount_rec
        BiffRecord.fnGroupCountRecord()
    end 
    def _datemode_rec
        BiffRecord.dateModeRecord(@dates_1904)        
    end
    def _precision_rec
        BiffRecord.precisionRecord(@use_cell_values)         
    end
    def _refresh_all_rec
        BiffRecord.refreshAllRecord()        
    end
    def _bookbool_rec
        BiffRecord.bookBoolRecord()         
    end
    def _all_fonts_num_formats_xf_styles_rec
        @styles.get_biff_data()
    end
    def _palette_rec
        ''
    end 
    def _useselfs_rec
        BiffRecord.useSelfsRecord()
    end
    def _boundsheets_rec(data_len_before, data_len_after, sheet_biff_lens)
        #  .................................  
        # BOUNDSEHEET0
        # BOUNDSEHEET1
        # BOUNDSEHEET2
        # ..................................
        # WORKSHEET0
        # WORKSHEET1
        # WORKSHEET2
        boundsheets_len = 0
        for sheet in @worksheets
            boundsheets_len += BiffRecord.boundSheetRecord(0x00, sheet.hidden, sheet.name).length
        end
        
        start = data_len_before + boundsheets_len + data_len_after
        
        result = ''
        for sheet_biff_len,  sheet in sheet_biff_lens.zip(@worksheets)
            result << BiffRecord.boundSheetRecord(start, sheet.hidden, sheet.name)
            start += sheet_biff_len
        end
        result
    end

    def _all_links_rec
        ''
    end
        
    def _sst_rec
        @sst.get_biff_record()
    end
        
    def _ext_sst_rec(abs_stream_pos)
        ''
        #BiffRecord.ExtSSTRecord(abs_stream_pos, @sst_record.str_placement,
        #@sst_record.portions_len)
    end

    def get_biff_data
        before = ''
        before << _bof_rec()
        before << _intf_hdr_rec()
        before << _intf_mms_rec()
        before << _intf_end_rec()
        before << _write_access_rec()
        before << _codepage_rec()
        before << _dsf_rec() 
        before << _tabid_rec() 
        before << _fngroupcount_rec()
        before << _wnd_protect_rec()
        before << _protect_rec()
        before << _obj_protect_rec()
        before << _password_rec()
        before << _prot4rev_rec()
        before << _prot4rev_pass_rec()
        before << _backup_rec()        
        before << _hide_obj_rec()        
        before << _window1_rec()
        before << _datemode_rec()
        before << _precision_rec()
        before << _refresh_all_rec()
        before << _bookbool_rec()
        before << _all_fonts_num_formats_xf_styles_rec()
        before << _palette_rec()
        before << _useselfs_rec()
         
        country            = _country_rec()
        all_links          = _all_links_rec()
        
        shared_str_table   = _sst_rec()
        after = country + all_links + shared_str_table
        
        ext_sst = _ext_sst_rec(0) # need fake cause we need calc stream pos
        eof = _eof_rec()
	@worksheets[@active_sheet].selected = true if @worksheets.length > 0
        sheets = ''
        sheet_biff_lens = []
        for sheet in @worksheets
            data = sheet.get_biff_data()
            sheets << data
            sheet_biff_lens << data.length
        end
        
        bundlesheets = _boundsheets_rec(before.length, after.length+ext_sst.length+eof.length, sheet_biff_lens)       
       
        sst_stream_pos = before.length + bundlesheets.length + country.length  + all_links.length
        ext_sst = _ext_sst_rec(sst_stream_pos)
        
        before + bundlesheets + after + ext_sst + eof + sheets
    end

    def save(file)
        doc = XlsDoc.new()
        doc.save(file, get_biff_data())
    end

    def binary
        doc = XlsDoc.new()
        doc.binary(get_biff_data())
    end
  end
end

if $0 == __FILE__
    wb = Excel::Workbook.new()
    wb.add_str('s')
    wb.add_str('sssssss')
    f = File.open('workbook1.bin', 'wb')
    f.write(wb.get_biff_data())
    f.close()
end