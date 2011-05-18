require File.dirname(__FILE__)+'/biff_records'
module Excel
  module BiffRecord
    def _position_image(sheet, row_start, col_start, x1, y1, width, height)
=begin
        Calculate the vertices that define the position of the image as required by
        the OBJ record.
    
                +------------+------------+
                |     A      |      B     |
          +-----+------------+------------+
          |     |(x1,y1)     |            |
          |  1  |(A1)._______|______      |
          |     |    |              |     |
          |     |    |              |     |
          +-----+----|    BITMAP    |-----+
          |     |    |              |     |
          |  2  |    |______________.     |
          |     |            |        (B2)|
          |     |            |     (x2,y2)|
          +---- +------------+------------+
    
        Example of a bitmap that covers some of the area from cell A1 to cell B2.
    
        Based on the width and height of the bitmap we need to calculate 8 vars
            col_start, row_start, col_end, row_end, x1, y1, x2, y2.
        The width and height of the cells are also variable and have to be taken into
        account.
        The values of col_start and row_start are passed in from the calling
        function. The values of col_end and row_end are calculated by subtracting
        the width and height of the bitmap from the width and height of the
        underlying cells.
        The vertices are expressed as a percentage of the underlying cell width as
        follows (rhs values are in pixels)
    
              x1 = X / W *1024
              y1 = Y / H *256
              x2 = (X-1) / W *1024
              y2 = (Y-1) / H *256
    
              Where=>  X is distance from the left side of the underlying cell
                      Y is distance from the top of the underlying cell
                      W is the width of the cell
                      H is the height of the cell
    
        Note=> the SDK incorrectly states that the height should be expressed as a
        percentage of 1024.
    
        col_start  - Col containing upper left corner of object
        row_start  - Row containing top left corner of object
        x1  - Distance to left side of object
        y1  - Distance to top of object
        width  - Width of image frame
        height  - Height of image frame
        
=end
        # Adjust start column for offsets that are greater than the col width
        while x1 >= sheet.col_width(col_start)
            x1 -= sheet.col_width(col_start)
            col_start += 1
        end
        # Adjust start row for offsets that are greater than the row height
        while y1 >= sheet.row_height(row_start)
            y1 -= sheet.row_height(row_start)
            row_start += 1
        end
        # Initialise end cell to the same as the start cell
        row_end = row_start   # Row containing bottom right corner of object
        col_end = col_start   # Col containing lower right corner of object
        width = width + x1 - 1
        height = height + y1 - 1
        # Subtract the underlying cell widths to find the end cell of the image
        while (width >= sheet.col_width(col_end))
            width -= sheet.col_width(col_end)
            col_end += 1
        end
        # Subtract the underlying cell heights to find the end cell of the image
        while (height >= sheet.row_height(row_end))
            height -= sheet.row_height(row_end)
            row_end += 1
        end
        # Bitmap isn't allowed to start or finish in a hidden cell, i.e. a cell
        # with zero height or width.
        if ((sheet.col_width( col_start) == 0) || (sheet.col_width( col_end) == 0) ||
            (sheet.row_height(row_start) == 0) || (sheet.row_height(row_end) == 0))
            return
        end
        # Convert the pixel values to the percentage value expected by Excel
        x1 = x1.to_f / sheet.col_width( col_start) * 1024
        y1 = y1.to_f / sheet.row_height(row_start) * 256
        # Distance to right side of object
        x2 = width.to_f  / sheet.col_width( col_end) * 1024
        # Distance to bottom of object
        y2 = height.to_f / sheet.row_height(row_end) * 256
        [col_start, x1, row_start, y1, col_end, x2, row_end, y2]
    end

    def objBmpRecord(row, col, sheet, im_data_bmp, x, y, scale_x, scale_y)

        # Scale the frame of the image.
        width = im_data_bmp.width * scale_x
        height = im_data_bmp.height * scale_y

        # Calculate the vertices of the image and write the OBJ record
        col_start, x1, row_start, y1, col_end, x2, row_end, y2 = _position_image(sheet, row, col, x, y, width, height)

=begin
            Store the OBJ record that precedes an IMDATA record. This could be generalise
            to support other Excel objects.
=end
        cObj = 0x0001      # Count of objects in file (set to 1)
        _OT = 0x0008        # Object type. 8 = Picture
        id = 0x0001        # Object ID
        grbit = 0x0614     # Option flags
        colL = col_start    # Col containing upper left corner of object
        dxL = x1            # Distance from left side of cell
        rwT = row_start     # Row containing top left corner of object
        dyT = y1            # Distance from top of cell
        colR = col_end      # Col containing lower right corner of object
        dxR = x2            # Distance from right of cell
        rwB = row_end       # Row containing bottom right corner of object
        dyB = y2            # Distance from bottom of cell
        cbMacro = 0x0000    # Length of FMLA structure
        reserved1 = 0x0000  # Reserved
        reserved2 = 0x0000  # Reserved
        icvBack = 0x09      # Background colour
        icvFore = 0x09      # Foreground colour
        fls = 0x00          # Fill pattern
        fAuto = 0x00        # Automatic fill
        icv = 0x08          # Line colour
        lns = 0xff          # Line style
        lnw = 0x01          # Line weight
        fAutoB = 0x00       # Automatic border
        frs = 0x0000        # Frame style
        cf = 0x0009         # Image format, 9 = bitmap
        reserved3 = 0x0000  # Reserved
        cbPictFmla = 0x0000 # Length of FMLA structure
        reserved4 = 0x0000  # Reserved
        grbit2 = 0x0001     # Option flags
        reserved5 = 0x0000  # Reserved
        
        get_biff_data(0x005D, [cObj, _OT, id, grbit, colL, dxL, rwT, dyT, colR, dxR, rwB,
                dyB, cbMacro, reserved1, reserved2, icvBack, icvFore, fls, fAuto,
                icv, lns, lnw, fAutoB, frs, cf, reserved3, cbPictFmla, Reserved4,
                grbit2, reserved5].pack('Vv12VvC8vVv4V'))
    end
    
    def _process_bitmap(bitmap)
=begin
        Convert a 24 bit bitmap into the modified internal format used by Windows.
        This is described in BITMAPCOREHEADER and BITMAPCOREINFO structures in the
        MSDN library.
=end
        # Open file and binmode the data in case the platform needs it.
        data = nil
        fh = File.open(bitmap, "rb") do |fh| data = fh.read() end
        # Check that the file is big enough to be a bitmap.
        if data.length <= 0x36
            raise Exception("bitmap doesn't contain enough data.")
        end
        # The first 2 bytes are used to identify the bitmap.
        if (data[0...(2)] != "BM")
            raise Exception("bitmap doesn't appear to to be a valid bitmap image.")
        end
        # Remove bitmap data=> ID.
        data = data[(2)..-1]
        # Read and remove the bitmap size. This is more reliable than reading
        # the data size at offset 0x22.
        #
        size = data[0...(4)].unpack("V")[0]
        size -=  0x36   # Subtract size of bitmap header.
        size +=  0x0C   # Add size of BIFF header.
        data = data[(4)..-1]
        # Remove bitmap data=> reserved, offset, header length.
        data = data[(12)..-1]
        # Read and remove the bitmap width and height. Verify the sizes.
        width, height = data[0...(8)].unpack("VV")
        data = data[(8)..-1]
        if (width > 0xFFFF)
            raise Exception("bitmap=> largest image width supported is 65k.")
        end
        if (height > 0xFFFF)
            raise Exception("bitmap=> largest image height supported is 65k.")
        end
        # Read and remove the bitmap planes and bpp data. Verify them.
        planes, bitcount = data[0...(4)].unpack("vv")
        data = data[(4)..-1]
        if (bitcount != 24)
            raise Exception("bitmap isn't a 24bit true color bitmap.")
        end
        if (planes != 1)
            raise Exception("bitmap=> only 1 plane supported in bitmap image.")
        end
        # Read and remove the bitmap compression. Verify compression.
        compression = data[0...(4)].unpack("V")[0]
        data = data[(4)..-1]
        if (compression != 0)
            raise Exception("bitmap=> compression not supported in bitmap image.")
        end
        # Remove bitmap data=> data size, hres, vres, colours, imp. colours.
        data = data[(20)..-1]
        # Add the BITMAPCOREHEADER data
        header = [0x000c, width, height, 0x01, 0x18].pack("Vvvvv")
        data = header + data
        [width, height, size, data]
    end
    
    def imDataBmpRecord
        rec_id  0x007F
=begin
            Insert a 24bit bitmap image in a worksheet. The main record required is
            IMDATA but it must be proceeded by a OBJ record to define its position.
=end
        @width, @height, @size, data = _process_bitmap(filename)
        # Write the IMDATA record to store the bitmap data
        cf = 0x09
        env = 0x01
        lcb = @size
        get_biff_data(0x007F, [cf, env, lcb, data].pack("vvVa*"))
    end
  end
end