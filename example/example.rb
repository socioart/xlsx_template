require "xlsx_template"

template = XlsxTemplate.load_file("#{__dir__}/template.xlsx")
template.render(
  "#{__dir__}/rendered.xlsx",
  name: "labocho",
  works: [
    {name: "foo", description: "Foo\nBar"},
    {name: "bar", description: "Foo\nBar"},
    {name: "baz", description: "Foo\nBar"},
  ],
  awards: [],
)
