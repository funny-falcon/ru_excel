# This implementation writes only 'Root Entry', 'Workbook' streams
# and 2 empty streams for aligning directory stream on sector boundary
# 
# LAYOUT
# 0         header
# 76                MSAT (1st part=> 109 SID)
# 512       workbook stream
# ...       additional MSAT sectors if streams' size > about 7 Mb == (109*512 * 128)
# ...       SAT
# ...       directory stream
#
# NOTE=> this layout is "ad hoc". It can be more general. RTFM
module Excel
  class XlsDoc
    SECTOR_SIZE = 0x0200
    MIN_LIMIT   = 0x1000

    SID_FREE_SECTOR  = -1
    SID_END_OF_CHAIN = -2
    SID_USED_BY_SAT  = -3
    SID_USED_BY_MSAT = -4

    def initialize
        #@book_stream = ''                # padded
        @book_stream_sect = []

        @dir_stream = ''
        @dir_stream_sect = []

        @packed_SAT = ''
        @SAT_sect = []

        @packed_MSAT_1st = ''
        @packed_MSAT_2nd = ''
        @MSAT_sect_2nd = []

        @header = ''
    end
    def self.__block(dentry_name, dentry_type, dentry_colour, dentry_did_left,
                     dentry_did_right, dentry_did_root, dentry_start_sid, dentry_stream_sz)
        [
           dentry_name,
           dentry_name.length,
           dentry_type,
           dentry_colour,
           dentry_did_left, 
           dentry_did_right,
           dentry_did_root,
           0, 0, 0, 0, 0, 0, 0, 0, 0,
           dentry_start_sid,
           dentry_stream_sz,
           0
        ].pack('a64 v C2 l3 L9 l L L')
    end
    STREAM_START = __block(
	"Root Entry\x00".gsub!(/(.)/,"\\1\x00"),
        0x05, # root storage
        0x01, # black
        -1,
        -1,
        1,
        -2,
        0)
    STREAM_WORKBOOK_DENTRY_NAME = "Workbook\x00".gsub!(/(.)/,"\\1\x00")
    STREAM_PADDING = __block(
        '',
        0x00, # empty
        0x01, # black
        -1,
        -1,
        -1,
        -2,
        0) * 2
    def _build_directory # align on sector boundary
        @dir_stream = ''
        @dir_stream << STREAM_START
        @dir_stream << XlsDoc.__block(
            STREAM_WORKBOOK_DENTRY_NAME,
            0x02, # user stream
            0x01, # black
            -1,
            -1,
            -1,
            0,
            @book_stream_len)

        @dir_stream << STREAM_PADDING
    end

    def _build_sat
        # Build SAT
        book_sect_count = @book_stream_len >> 9
        dir_sect_count  = @dir_stream.length >> 9
        
        total_sect_count     = book_sect_count + dir_sect_count
        _SAT_sect_count       = 0
        _MSAT_sect_count      = 0
        _SAT_sect_count_limit = 109
        while total_sect_count > 128*_SAT_sect_count or _SAT_sect_count > _SAT_sect_count_limit
            _SAT_sect_count  += 1
            total_sect_count += 1
            if _SAT_sect_count > _SAT_sect_count_limit
                _MSAT_sect_count      += 1
                total_sect_count     += 1
                _SAT_sect_count_limit += 127
            end
        end


        _SAT = [SID_FREE_SECTOR]*128*_SAT_sect_count

        sect = 0
        while sect < book_sect_count - 1
            @book_stream_sect << sect
            _SAT[sect] = sect + 1
            sect += 1
        end
        @book_stream_sect << sect
        _SAT[sect] = SID_END_OF_CHAIN
        sect += 1

        while sect < book_sect_count + _MSAT_sect_count
            @MSAT_sect_2nd << sect
            _SAT[sect] = SID_USED_BY_MSAT
            sect += 1
        end

        while sect < book_sect_count + _MSAT_sect_count + _SAT_sect_count
            @SAT_sect << sect
            _SAT[sect] = SID_USED_BY_SAT
            sect += 1
        end

        while sect < book_sect_count + _MSAT_sect_count + _SAT_sect_count + dir_sect_count - 1
            @dir_stream_sect << sect
            _SAT[sect] = sect + 1
            sect += 1
        end
        @dir_stream_sect << sect
        _SAT[sect] = SID_END_OF_CHAIN
        sect += 1

        @packed_SAT = _SAT.pack('l%d'%(_SAT_sect_count*128))

	
        _MSAT_1st = @SAT_sect + [SID_FREE_SECTOR]*([109-@SAT_sect.length,0].max)

        @packed_MSAT_1st = _MSAT_1st.pack('l109')

        _MSAT_2nd = [SID_FREE_SECTOR]*128*_MSAT_sect_count
        if _MSAT_sect_count > 0
            _MSAT_2nd[- 1] = SID_END_OF_CHAIN
        end

        i = 109
        msat_sect = 0
        sid_num = 0
        while i < _SAT_sect_count
            if (sid_num + 1) % 128 == 0
                #print 'link: ',
                msat_sect += 1
                if msat_sect < @MSAT_sect_2nd.length
                    _MSAT_2nd[sid_num] = @MSAT_sect_2nd[msat_sect]
                end
            else
                #print 'sid: ',
                _MSAT_2nd[sid_num] = @SAT_sect[i]
                i += 1
            end
            #print sid_num, MSAT_2nd[sid_num]
            sid_num += 1
        end

        @packed_MSAT_2nd = _MSAT_2nd.pack('l%d'%(_MSAT_sect_count*128))

        #print vars()
        #print zip(range(0, sect), SAT)
        #print @book_stream_sect
        #print self.MSAT_sect_2nd
        #print MSAT_2nd
        #print self.SAT_sect
        #print @dir_stream_sect
    end

    def _build_header
        doc_magic             = "\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"
        file_uid              = "\x00"*16
        rev_num               = "\x3E\x00"
        ver_num               = "\x03\x00"
        byte_order            = "\xFE\xFF"
        log_sect_size         = [ 9].pack('v')
        log_short_sect_size   = [ 6].pack('v')
        not_used0             = "\x00"*10
        total_sat_sectors     = [@SAT_sect.length].pack('V')
        dir_start_sid         = [ @dir_stream_sect[0]].pack('V')
        not_used1             = "\x00"*4
        min_stream_size       = [ 0x1000].pack('V')
        ssat_start_sid        = [ -2].pack('V')
        total_ssat_sectors    = [ 0].pack('V')

        if @MSAT_sect_2nd.length == 0
            msat_start_sid        = [ -2].pack('V')
        else
            msat_start_sid        = [@MSAT_sect_2nd[0]].pack('V')
        end

        total_msat_sectors    = [@MSAT_sect_2nd.length].pack('V')

        @header =                    [  doc_magic,
                                        file_uid,
                                        rev_num,
                                        ver_num,
                                        byte_order,
                                        log_sect_size,
                                        log_short_sect_size,
                                        not_used0,
                                        total_sat_sectors,
                                        dir_start_sid,
                                        not_used1,
                                        min_stream_size,
                                        ssat_start_sid,
                                        total_ssat_sectors,
                                        msat_start_sid,
                                        total_msat_sectors
                                    ].join('')
    end

    def save(f, stream)
        # 1. Align stream on 0x1000 boundary (and therefore on sector boundary)
	is_string = f.is_a? String
	f = File.new(f, 'wb') if is_string
        padding = "\x00" * (0x1000 - (stream.length % 0x1000))
        @book_stream_len = stream.length + padding.length

        _build_directory()
        _build_sat()
        _build_header()
        
        f.write(@header)
        f.write(@packed_MSAT_1st)
        f.write(stream)
        f.write(padding)
        f.write(@packed_MSAT_2nd)
        f.write(@packed_SAT)
        f.write(@dir_stream)
	f.close if is_string
    end

    def binary(stream)
        padding = "\x00" * (0x1000 - (stream.length % 0x1000))
        @book_stream_len = stream.length + padding.length

        _build_directory()
        _build_sat()
        _build_header()
        "#@header#@packed_MSAT_1st#{stream}#{padding}#@packed_MSAT_2nd#@packed_SAT#@dir_stream"
    end
  end
end

if $0 == __FILE__
    d = XlsDoc()
    d.save('a.aaa', 'b'*17000)
end




