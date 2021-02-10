# XlsxTemplate

Create .xlsx from .xlsx file with Handlebars like syntax.

## Installation

Add this line to your application's Gemfile:

```ruby
gem 'xlsx_template', git: "https://github.com/socioart/xlsx_template.git"
```

And then execute:

    $ bundle install

## Limitation

* Does not workbook with any formula.

## Usage

See `example` directory.

### beforeRow

If `A1` cell value is `beforeRow`, column `A` has been interpreted specially.

- `{{#if TOP_LEVEL_VARIABLE_NAME}}` If value of `TOP_LEVEL_VARIABLE_NAME` is *not* truthy (`nil, false, "", 0, or []`), remove row.
- `{{#each TOP_LEVEL_VARIABLE_NAME}}` Repeats row for value of `TOP_LEVEL_VARIABLE_NAME`.

In the output file, values of column `A` are removed and the column is hidden.

## Development

After checking out the repo, run `bin/setup` to install dependencies. Then, run `rake spec` to run the tests. You can also run `bin/console` for an interactive prompt that will allow you to experiment.

To install this gem onto your local machine, run `bundle exec rake install`. To release a new version, update the version number in `version.rb`, and then run `bundle exec rake release`, which will create a git tag for the version, push git commits and tags, and push the `.gem` file to [rubygems.org](https://rubygems.org).

## Contributing

Bug reports and pull requests are welcome on GitHub at https://github.com/labocho/xlsx_template.


## License

The gem is available as open source under the terms of the [MIT License](https://opensource.org/licenses/MIT).
