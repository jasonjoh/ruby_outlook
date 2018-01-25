# -*- encoding: utf-8 -*-
# stub: ruby_outlook 0.1.0 ruby lib

Gem::Specification.new do |s|
  s.name = "ruby_outlook"
  s.version = "0.1.0"

  s.required_rubygems_version = Gem::Requirement.new(">= 0") if s.respond_to? :required_rubygems_version=
  s.metadata = { "allowed_push_host" => "TODO: Set to 'http://mygemserver.com' to prevent pushes to rubygems.org, or delete to allow pushes to any server." } if s.respond_to? :metadata=
  s.require_paths = ["lib"]
  s.authors = ["Jason Johnston"]
  s.bindir = "exe"
  s.date = "2016-04-25"
  s.description = "This ruby gem provides functions for common operations with the Office 365 Mail, Calendar, and Contacts APIs."
  s.email = ["jasonjoh@microsoft.com"]
  s.files = [".gitattributes", ".gitignore", ".travis.yml", "Gemfile", "LICENSE.TXT", "README.md", "Rakefile", "bin/console", "bin/setup", "lib/ruby_outlook.rb", "lib/ruby_outlook/version.rb", "lib/run-tests.rb", "ruby_outlook.gemspec"]
  s.homepage = "https://github.com/jasonjoh/ruby_outlook"
  s.rubygems_version = "2.5.1"
  s.summary = "A ruby gem to invoke the Office 365 REST APIs."

  if s.respond_to? :specification_version then
    s.specification_version = 4

    if Gem::Version.new(Gem::VERSION) >= Gem::Version.new('1.2.0') then
      s.add_runtime_dependency(%q<faraday>, [">= 0"])
      s.add_runtime_dependency(%q<uuidtools>, [">= 0"])
      s.add_development_dependency(%q<bundler>, ["~> 1.8"])
      s.add_development_dependency(%q<rake>, ["~> 10.0"])
    else
      s.add_dependency(%q<faraday>, [">= 0"])
      s.add_dependency(%q<uuidtools>, [">= 0"])
      s.add_dependency(%q<bundler>, ["~> 1.8"])
      s.add_dependency(%q<rake>, ["~> 10.0"])
    end
  else
    s.add_dependency(%q<faraday>, [">= 0"])
    s.add_dependency(%q<uuidtools>, [">= 0"])
    s.add_dependency(%q<bundler>, ["~> 1.8"])
    s.add_dependency(%q<rake>, ["~> 10.0"])
  end
end
