# encoding: utf-8
require "xlsxtream/xml"
require "xlsxtream/row"

module Xlsxtream
  class Worksheet
    def initialize(io, sst = nil, cols_width = nil)
      @io = io
      @rownum = 1
      @sst = sst
      write_header(cols_width)
    end

    def <<(row)
      @io << Row.new(row, @rownum, @sst).to_xml
      @rownum += 1
    end
    alias_method :add_row, :<<

    def close
      write_footer
    end

    private

    def write_header(cols_width)
      @io << XML.header
      @io << XML.strip(<<-XML)
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      XML

      # for the width
      write_cols(cols_width) unless cols_width.nil?

      @io << XML.strip(<<-XML)
          <sheetData>
      XML
    end

    def write_cols(cols_width)
      cols = '<cols>'
      cols_width.each_with_index do |e, i|
        cols += "<col min=\"#{i + 1}\" max=\"#{i + 1}\" bestFit=\"1\" customWidth=\"1\" width=\"#{e}\"/>"
      end
      cols += '</cols>'
      @io << XML.strip(cols)
    end

    def write_footer
      @io << XML.strip(<<-XML)
          </sheetData>
        </worksheet>
      XML
    end
  end
end
