# encoding: UTF-8

require "docx_replace/version"
require 'zip'
require 'tempfile'

module DocxReplace
  class Doc
    attr_reader :document_content

    IO_METHODS = [:tell, :seek, :read, :close].freeze
    DOCUMENT_FILE_PATH = 'word/document.xml'.freeze

    def initialize(path_or_io, temp_dir=nil)
      if IO_METHODS.all? { |method| path_or_io.respond_to? method }
        path_or_io.seek 0
        Zip::File.open_buffer(path_or_io) do |zf|
          @zip_file = zf
        end
      else
        @zip_file = Zip::File.new(path_or_io)
      end
      @temp_dir = temp_dir

      read_docx_file
    end

    def replace(pattern, replacement, multiple_occurrences=false)
      replace = replacement.to_s
      if multiple_occurrences
        @document_content.force_encoding('UTF-8').gsub!(pattern, replace)
      else
        @document_content.force_encoding('UTF-8').sub!(pattern, replace)
      end
    end

    def matches(pattern)
      @document_content.scan(pattern).map{|match| match.first}
    end

    def unique_matches(pattern)
      matches(pattern)
    end

    alias_method :uniq_matches, :unique_matches

    def commit(new_path=nil)
      write_back_to_file(new_path)
    end

    def commit_to_stream(io)
      io.binmode if io.respond_to? :binmode
      r = Zip::OutputStream.write_buffer io do |zos|
        @zip_file.entries.each do |entry|
          zos.put_next_entry(entry.name)
          if entry.name == DOCUMENT_FILE_PATH
            zos.write @document_content
            else
            zos.copy_raw_entry entry
          end
        end
      end
      r.flush
      io.reopen r
      io.rewind
      io
    end

    private

    def read_docx_file
      @document_content = @zip_file.read(DOCUMENT_FILE_PATH)
    end

    def write_back_to_file(new_path=nil)
      if @temp_dir.nil?
        temp_file = Tempfile.new('docxedit-')
      else
        temp_file = Tempfile.new('docxedit-', @temp_dir)
      end

      self.commit_to_stream(temp_file)

      if new_path.nil?
        path = @zip_file.name
        FileUtils.rm(path)
      else
        path = new_path
      end
      FileUtils.mv(temp_file.path, path)
      @zip_file = Zip::File.new(path)
    end
  end
end
