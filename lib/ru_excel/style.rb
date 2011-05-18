dir = File.dirname(__FILE__)
require dir+'/formatting'
require dir+'/biff_records'

module Excel

  class XFStyle
    Default_num_format = 'general'
    Default_font = Font.new()
    Default_alignment = Alignment.new()
    Default_borders = Borders.new()
    Default_pattern = Pattern.new()
    Default_protection = Protection.new()

    def initialize
        @num_format_str  = Default_num_format
        @font            = Default_font 
        @alignment       = Default_alignment
        @borders         = Default_borders
        @pattern         = Default_pattern 
        @protection      = Default_protection
    end
    attr_accessor :num_format_str, :font, :alignment, :borders, :pattern, :protection
  end

  class StyleCollection
    Std_num_fmt_list = [
            'general',
            '0',
            '0.00',
            '#,##0',
            '#,##0.00',
            '"$"#,##0_);("$"#,##',
            '"$"#,##0_);[Red]("$"#,##',
            '"$"#,##0.00_);("$"#,##',
            '"$"#,##0.00_);[Red]("$"#,##',
            '0%',
            '0.00%',
            '0.00E+00',
            '# ?/?',
            '# ??/??',
            'M/D/YY',
            'D-MMM-YY',
            'D-MMM',
            'MMM-YY',
            'h:mm AM/PM',
            'h:mm:ss AM/PM',
            'h:mm',
            'h:mm:ss',
            'M/D/YY h:mm',
            '_(#,##0_);(#,##0)',
            '_(#,##0_);[Red](#,##0)',
            '_(#,##0.00_);(#,##0.00)',
            '_(#,##0.00_);[Red](#,##0.00)',
            '_("$"* #,##0_);_("$"* (#,##0);_("$"* "-"_);_(@_)',
            '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)',
            '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)',
            '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)',
            'mm:ss',
            '[h]:mm:ss',
            'mm:ss.0',
            '##0.0E+0',
            '@'   
    ]

    NumFormats = {}
    for fmtidx, fmtstr in (0...23).zip(Std_num_fmt_list[(0)...(23)])
        NumFormats[fmtstr] = fmtidx
    end
    for fmtidx, fmtstr in (37...50).zip(Std_num_fmt_list[(23)..-1])
        NumFormats[fmtstr] = fmtidx 
    end

    def initialize
        @_fonts = {}
        @_fonts[Font.new()] = 0
        @_fonts[Font.new()] = 1
        @_fonts[Font.new()] = 2
        @_fonts[Font.new()] = 3
        # The font with index 4 is omitted in all BIFF versions
        @_fonts[Font.new()] = 5
        @_num_formats = NumFormats.dup

        @_xf = {}
        @default_style = XFStyle.new()
        @_default_xf = _add_style(@default_style)[0]
    end
    attr_accessor :default_style

    def add(style)
        style == nil ? 0x10 : _add_style(style)[1]
    end

    def _add_style(style)
        num_format_str = style.num_format_str
        num_format_idx = @_num_formats[num_format_str] ||=
            163 + @_num_formats.length - Std_num_fmt_list.length
        font = style.font
        font_idx = @_fonts[font] ||= @_fonts.length + 1
        xf = [font_idx, num_format_idx, style.alignment, style.borders, style.pattern, style.protection]
        xf_index = @_xf[xf] ||= 0x10 + @_xf.length
        [xf, xf_index]
    end

    def get_biff_data
        result = ''
        result << _all_fonts()
        result << _all_num_formats()
        result << _all_cell_styles()
        result << _all_styles()
        result
    end

    def _all_fonts
        fonts = @_fonts.map{|k,v| [v,k]}.sort!
        fonts.collect!{|_,font| font.get_biff_record}
        fonts.join('')
    end

    def _all_num_formats
        formats = @_num_formats.select{|k, v| v>=163}.to_a.each{|a| a.reverse!}
        formats.map!{|fmtidx, fmtstr| BiffRecord.numberFormatRecord(fmtidx, fmtstr)}
        formats.join('')
    end

    def _all_cell_styles
        result = BiffRecord.xfRecord(@_default_xf, 'style') * 16
        result << @_xf.map{|k,v| [v,k]}.sort!.collect!{|_,xf| BiffRecord.xfRecord(xf)}.join('')
    end

    def _all_styles
        BiffRecord.styleRecord()
    end
  end
end
