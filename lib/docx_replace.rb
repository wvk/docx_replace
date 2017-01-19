# encoding: UTF-8

require 'docx_replace/version'
require 'zip'
require 'tempfile'

module DocxReplace
  class Doc
    attr_reader :document_content

    IO_METHODS = [:tell, :seek, :read, :close].freeze

    def initialize(path_or_io, temp_dir=nil)
      if IO_METHODS.all? { |method| path_or_io.respond_to? method }
        path_or_io.seek 0
        Zip::File.open_buffer(path_or_io) do |zf|
          @zip_file = zf
        end
      else
        @zip_file = Zip::File.new(path_or_io)
      end
      @document_file_paths = find_query_file_paths(@zip_file)

      @temp_dir = temp_dir
      read_docx_files
    end

    def replace(pattern, replacement, multiple_occurrences = false)
      replacement = CGI.escapeHTML(replacement.to_s.encode(xml: :text)).gsub("\n", '</w:t><w:br/><w:t>').gsub("\t", '</w:t><w:tab/><w:t>')
      @document_contents.each do |path, document_content|
        if multiple_occurrences
          document_content.force_encoding('UTF-8').gsub!(pattern, replacement)
        else
          document_content.force_encoding('UTF-8').sub!(pattern, replacement)
        end
      end
    end

    def matches(pattern)
      @document_contents.values.join.scan(pattern).map{|match| match.first}
    end

    def unique_matches(pattern)
      matches(pattern)
    end

    alias_method :uniq_matches, :unique_matches

    def commit(new_path=nil)
      write_back_to_file(new_path)
    end

    def commit_to_stream(io = ::StringIO.new(''))
      io.binmode if io.respond_to? :binmode
      r = Zip::OutputStream.write_buffer io do |zos|
         @zip_file.entries.each do |e|
          unless @document_file_paths.include?(e.name)
            zos.put_next_entry e.name
            zos.print e.get_input_stream.read
          end
        end

        @document_contents.each do |path, document|
          zos.put_next_entry path
          zos.print document
        end

      end
      r.flush
      io.reopen r
      io.rewind
      io
    end

    private

    def find_query_file_paths(zipfile)
      zipfile.entries.map(&:name).select do |entry|
        !(/^word\/(document|footer[0-9]+|header[0-9]+).xml$/ =~ entry).nil?
      end
    end

    def read_docx_files
      @document_contents = {}
      @document_file_paths.each do |path|
        @document_contents[path] = @zip_file.read(path)
      end
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
