require_relative "lib/xlsx_template/version"

Gem::Specification.new do |spec|
  spec.name          = "xlsx_template"
  spec.version       = XlsxTemplate::VERSION
  spec.authors       = ["labocho"]
  spec.email         = ["labocho@penguinlab.jp"]

  spec.summary       = "Create .xlsx from .xlsx file with Handlebars like syntax"
  spec.description   = "Create .xlsx from .xlsx file with Handlebars like syntax"
  spec.homepage      = "https://github.com/socioart/xlsx_template"
  spec.license       = "MIT"
  spec.required_ruby_version = Gem::Requirement.new(">= 2.4.0")

  # spec.metadata["allowed_push_host"] = "TODO: Set to 'http://mygemserver.com'"

  spec.metadata["homepage_uri"] = spec.homepage
  spec.metadata["source_code_uri"] = "https://github.com/socioart/xlsx_template"
  spec.metadata["changelog_uri"] = "https://github.com/socioart/xlsx_template/blob/master/CHANGELOG.md"

  # Specify which files should be added to the gem when it is released.
  # The `git ls-files -z` loads the files in the RubyGem that have been added into git.
  spec.files = Dir.chdir(File.expand_path(__dir__)) do
    `git ls-files -z`.split("\x0").reject {|f| f.match(%r{^(test|spec|features)/}) }
  end
  spec.bindir        = "exe"
  spec.executables   = spec.files.grep(%r(^exe/)) {|f| File.basename(f) }
  spec.require_paths = ["lib"]

  spec.add_dependency "ruby-handlebars", "~> 0.4.0"
  spec.add_dependency "rubyXL", "~> 3.4.17"
end
