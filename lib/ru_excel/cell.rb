require File.dirname(__FILE__)+'/biff_records'

module Excel
  module Cell
    extend self
    def strCell(parent, idx, xf_idx, sst_idx)
        BiffRecord.labelSSTRecord(parent.get_index, idx, xf_idx, sst_idx)
    end
    def blankCell(parent, idx, xf_idx)
        BiffRecord.blankRecord(parent.get_index, idx, xf_idx)
    end
    def mulBlankCell(parent, col1, col2, xf_idx)
        BiffRecord.mulBlankRecord(parent.get_index, col1, col2, xf_idx)
    end
    def numberCell(parent, idx, xf_idx, number)
        rk_encoded = 0
        packed = [number.to_f].pack('E')

        #print @number
        w0, w1, w2, w3 = packed.unpack('v4')
        if w0 == 0 && w1 == 0 && w2 & 0xFFFC == w2
            # 34 lsb are 0
            #print "float RK"
            rk_encoded = (w3 << 16) | w2
            BiffRecord.rkRecord(parent.get_index, idx, xf_idx, rk_encoded)
        elsif number.abs < 0x40000000 && number.to_i == number
            #print "30-bit integer RK"
            rk_encoded = 2 | (number.to_i << 2)
            BiffRecord.rkRecord(parent.get_index, idx, xf_idx, rk_encoded)
        else
          temp = number*100
          packed100 = [ temp].pack('E')
          w0, w1, w2, w3 = packed100.unpack('v4')
          if w0 == 0 && w1 == 0 && w2 & 0xFFFC == w2
              # 34 lsb are 0
              #print "float RK*100"
              rk_encoded = 1 | (w3 << 16) | w2
              BiffRecord.rkRecord(parent.get_index, idx, xf_idx, rk_encoded)
          elsif temp.abs < 0x40000000 && temp.to_i == temp
              #print "30-bit integer RK*100"
              rk_encoded = 3 | (temp.to_i << 2)
              BiffRecord.rkRecord(parent.get_index, idx, xf_idx, rk_encoded)
          else
          #print "Number" 
          #print
              BiffRecord.numberRecord(parent.get_index, idx, xf_idx, number)
          end
        end
    end
    def mulNumberCell(parent, idx, xf_idx, sst_idx)
        raise Exception
    end
    def fomulaCell(parent, idx, xf_idx, frmla)
        BiffRecords.formulaRecord(parent.get_index, idx, xf_idx, frmla.rpn())
    end
  end
end