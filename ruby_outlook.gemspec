# coding: utf-8
lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require 'ruby_outlook/version'

Gem::Specification.new do |spec|
  spec.name          = "ruby_outlook"
  spec.version       = RubyOutlook::VERSION
  spec.authors       = ["Jason Johnston"]
  spec.email         = ["jasonjoh@microsoft.com"]

  spec.summary       = %q{A ruby gem to invoke the Outlook REST APIs.}
  spec.description   = %q{This ruby gem provides functions for common operations with the Outlook Mail, Calendar, and Contacts APIs.}
  spec.homepage      = "https://github.com/jasonjoh/ruby_outlook"

  spec.files         = `git ls-files -z`.split("\x0").reject { |f| f.match(%r{^(test|spec|features)/}) }
  spec.bindir        = "exe"
  spec.executables   = spec.files.grep(%r{^exe/}) { |f| File.basename(f) }
  spec.require_paths = ["lib"]
  
  spec.add_dependency "faraday"

  spec.add_development_dependency "bundler", "~> 1.8"
  spec.add_development_dependency "rake", "~> 10.0"
end
