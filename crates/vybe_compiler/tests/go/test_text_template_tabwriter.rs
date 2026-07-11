//! text/tabwriter and text/template compile patterns.

go_compile_cases! {
    tabwriter_new => "package main; import \"text/tabwriter\"; func main() { w := tabwriter.NewWriter(nil, 0, 0, 1, ' ', 0); _ = w }",
    tabwriter_write => "package main; import \"text/tabwriter\"; import \"os\"; func main() { w := tabwriter.NewWriter(os.Stdout, 0, 0, 1, ' ', 0); _, _ = w.Write([]byte(\"a\\tb\\n\")) }",
    tabwriter_flush => "package main; import \"text/tabwriter\"; import \"bytes\"; func main() { var b bytes.Buffer; w := tabwriter.NewWriter(&b, 0, 0, 1, ' ', 0); w.Flush() }",
    template_parse => "package main; import \"text/template\"; func main() { _, _ = template.New(\"t\").Parse(\"{{.}}\") }",
    template_execute => "package main; import \"text/template\"; import \"bytes\"; func main() { t, _ := template.New(\"t\").Parse(\"{{.}}\"); var b bytes.Buffer; _ = t.Execute(&b, \"hi\") }",
    template_must => "package main; import \"text/template\"; func main() { _ = template.Must(template.New(\"t\").Parse(\"x\")) }",
    template_func_map => "package main; import \"text/template\"; import \"strings\"; func main() { _, _ = template.New(\"t\").Funcs(template.FuncMap{\"U\": strings.ToUpper}).Parse(\"{{U .}}\") }",
    template_parse_files => "package main; import \"text/template\"; func main() { _, _ = template.ParseFiles(\"tmpl.txt\") }",
    template_parse_glob => "package main; import \"text/template\"; func main() { _, _ = template.ParseGlob(\"*.tmpl\") }",
    template_define => "package main; import \"text/template\"; func main() { t := template.New(\"root\"); _ = t.New(\"child\") }",
}
