require "rubygems"
require "bundler/setup"
Bundler.require(:default)

unless Kernel.respond_to?(:require_relative)
  module Kernel
    def require_relative(path)
      require File.join(File.dirname(caller[0]), path.to_str)
    end
  end
end

require_relative "docxedit/docx"
