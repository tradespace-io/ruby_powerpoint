require 'zip/filesystem'
require 'nokogiri'

module RubyPowerpoint

  class RubyPowerpoint::Presentation

    attr_reader :files

    def initialize path_or_io
      if path_or_io.is_a?(StringIO)
        @files = Zip::File.open_buffer path_or_io
      else
        raise 'Not a valid file format.' unless (['.pptx'].include? File.extname(path_or_io).downcase)
        @files = Zip::File.open path_or_io
      end
    end

    def slides
      slides = Array.new
      @files.each do |f|
        if f.name.include? 'ppt/slides/slide'
          slides.push RubyPowerpoint::Slide.new(self, f.name)
        end
      end
      slides.sort{|a,b| a.slide_num <=> b.slide_num}
    end

    def close
      @files.close
    end
  end
end
