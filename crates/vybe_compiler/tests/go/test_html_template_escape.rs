//! html/template escaping: Parse/Execute with auto-escape, html.EscapeString,
//! template JS/URL/CSS typed values, MustParse — distinct from compile smokes in
//! `test_cover_text_html_log.rs` and `test_stdlib_mime_runtime.rs`.


go_run_cases! {
    html_escape_string_angle_brackets => (
        "package main; import \"fmt\"; import \"html\"; func main() { fmt.Println(html.EscapeString(\"<script>\")) }",
        vec!["&lt;script&gt;"]
    ),
    html_escape_string_ampersand => (
        "package main; import \"fmt\"; import \"html\"; func main() { fmt.Println(html.EscapeString(\"a&b\")) }",
        vec!["a&amp;b"]
    ),
    html_escape_string_quotes => (
        "package main; import \"fmt\"; import \"html\"; func main() { fmt.Println(html.EscapeString(`\"hi\"`)) }",
        vec!["&#34;hi&#34;"]
    ),
    html_escape_string_plain_text => (
        "package main; import \"fmt\"; import \"html\"; func main() { fmt.Println(html.EscapeString(\"hello\")) }",
        vec!["hello"]
    ),
    html_unescape_string_entities => (
        "package main; import \"fmt\"; import \"html\"; func main() { fmt.Println(html.UnescapeString(\"&lt;b&gt;\")) }",
        vec!["<b>"]
    ),
    html_template_execute_escapes_html => (
        "package main; import \"fmt\"; import \"html/template\"; import \"bytes\"; func main() { t := template.Must(template.New(\"p\").Parse(\"{{.}}\")); var b bytes.Buffer; t.Execute(&b, \"<b>\"); fmt.Println(b.String()) }",
        vec!["&lt;b&gt;"]
    ),
    html_template_execute_preserves_safe_text => (
        "package main; import \"fmt\"; import \"html/template\"; import \"bytes\"; func main() { t := template.Must(template.New(\"p\").Parse(\"Hello {{.}}\")); var b bytes.Buffer; t.Execute(&b, \"World\"); fmt.Println(b.String()) }",
        vec!["Hello World"]
    ),
    html_template_execute_struct_field => (
        "package main; import \"fmt\"; import \"html/template\"; import \"bytes\"; func main() { type Page struct { Title string }; t := template.Must(template.New(\"p\").Parse(\"{{.Title}}\")); var b bytes.Buffer; t.Execute(&b, Page{Title: \"<x>\"}); fmt.Println(b.String()) }",
        vec!["&lt;x&gt;"]
    ),
    html_template_execute_escaped_pipeline => (
        "package main; import \"fmt\"; import \"html/template\"; import \"bytes\"; func main() { t := template.Must(template.New(\"p\").Parse(`{{. | html}}`)); var b bytes.Buffer; t.Execute(&b, \"<a>\"); fmt.Println(b.String()) }",
        vec!["&lt;a&gt;"]
    ),
    html_template_execute_with_no_escape => (
        "package main; import \"fmt\"; import \"html/template\"; import \"bytes\"; func main() { t := template.Must(template.New(\"p\").Parse(`{{.}}`)); var b bytes.Buffer; t.Execute(&b, template.HTML(\"<b>ok</b>\")); fmt.Println(b.String()) }",
        vec!["<b>ok</b>"]
    ),
    html_template_js_type_value => (
        "package main; import \"fmt\"; import \"html/template\"; import \"bytes\"; func main() { t := template.Must(template.New(\"p\").Parse(`{{.}}`)); var b bytes.Buffer; t.Execute(&b, template.JS(\"alert(1)\")); fmt.Println(len(b.String()) > 0) }",
        vec!["true"]
    ),
    html_template_url_type_value => (
        "package main; import \"fmt\"; import \"html/template\"; import \"bytes\"; func main() { t := template.Must(template.New(\"p\").Parse(\"{{.}}\")); var b bytes.Buffer; t.Execute(&b, template.URL(\"http://example.com\")); fmt.Println(len(b.String()) > 10) }",
        vec!["true"]
    ),
    html_template_css_type_value => (
        "package main; import \"fmt\"; import \"html/template\"; import \"bytes\"; func main() { t := template.Must(template.New(\"p\").Parse(`{{.}}`)); var b bytes.Buffer; t.Execute(&b, template.CSS(\"color: red\")); fmt.Println(len(b.String()) > 0) }",
        vec!["true"]
    ),
    html_template_html_type_bypasses_escape => (
        "package main; import \"fmt\"; import \"html/template\"; import \"bytes\"; func main() { t := template.Must(template.New(\"p\").Parse(`{{.}}`)); var b bytes.Buffer; t.Execute(&b, template.HTML(\"<em>hi</em>\")); fmt.Println(b.String()) }",
        vec!["<em>hi</em>"]
    ),
    html_template_js_escape_string => (
        "package main; import \"fmt\"; import \"html/template\"; func main() { fmt.Println(template.JSEscapeString(`\"`)) }",
        vec!["\\u0022"]
    ),
    html_template_url_escape_string => (
        "package main; import \"fmt\"; import \"html/template\"; func main() { s := template.URLEscapeString(\"a b\"); fmt.Println(len(s) > 0) }",
        vec!["true"]
    ),
    html_template_css_escape_string => (
        "package main; import \"fmt\"; import \"html/template\"; func main() { s := template.CSSEscapeString(\"<style>\"); fmt.Println(len(s) > 0) }",
        vec!["true"]
    ),
    html_template_html_escape_string => (
        "package main; import \"fmt\"; import \"html/template\"; func main() { fmt.Println(template.HTMLEscapeString(\"<div>\")) }",
        vec!["&lt;div&gt;"]
    ),
    html_template_parse_and_execute => (
        "package main; import \"fmt\"; import \"html/template\"; import \"bytes\"; func main() { t, _ := template.New(\"x\").Parse(\"{{.N}}\"); var b bytes.Buffer; t.Execute(&b, struct{ N int }{7}); fmt.Println(b.String()) }",
        vec!["7"]
    ),
    html_template_execute_template_named => (
        "package main; import \"fmt\"; import \"html/template\"; import \"bytes\"; func main() { t := template.Must(template.New(\"base\").Parse(\"{{.}}\")); var b bytes.Buffer; t.ExecuteTemplate(&b, \"base\", \"ok\"); fmt.Println(b.String()) }",
        vec!["ok"]
    ),
    html_template_conditional_escape => (
        "package main; import \"fmt\"; import \"html/template\"; import \"bytes\"; func main() { t := template.Must(template.New(\"p\").Parse(`{{if .}}yes{{else}}no{{end}}`)); var b bytes.Buffer; t.Execute(&b, true); fmt.Println(b.String()) }",
        vec!["yes"]
    ),
    html_template_range_escape => (
        "package main; import \"fmt\"; import \"html/template\"; import \"bytes\"; func main() { t := template.Must(template.New(\"p\").Parse(`{{range .}}{{.}}{{end}}`)); var b bytes.Buffer; t.Execute(&b, []string{\"<a>\", \"b\"}); fmt.Println(b.String()) }",
        vec!["&lt;a&gt;b"]
    ),
    html_escape_string_less_than => (
        "package main; import \"fmt\"; import \"html\"; func main() { fmt.Println(html.EscapeString(\"1<2\")) }",
        vec!["1&lt;2"]
    ),
    html_escape_string_greater_than => (
        "package main; import \"fmt\"; import \"html\"; func main() { fmt.Println(html.EscapeString(\"2>1\")) }",
        vec!["2&gt;1"]
    ),
    html_template_attr_escape => (
        "package main; import \"fmt\"; import \"html/template\"; import \"bytes\"; func main() { t := template.Must(template.New(\"p\").Parse(`<a title=\"{{.}}\">x</a>`)); var b bytes.Buffer; t.Execute(&b, `say \"hi\"`); fmt.Println(len(b.String()) > 5) }",
        vec!["true"]
    ),
    html_template_url_query_escaper => (
        "package main; import \"fmt\"; import \"html/template\"; func main() { s := template.URLQueryEscaper([]byte(\"q=go lang\")); fmt.Println(len(s) > 0) }",
        vec!["true"]
    ),
    html_template_execute_nil_data => (
        "package main; import \"fmt\"; import \"html/template\"; import \"bytes\"; func main() { t := template.Must(template.New(\"p\").Parse(\"static\")); var b bytes.Buffer; t.Execute(&b, nil); fmt.Println(b.String()) }",
        vec!["static"]
    ),
    html_template_nested_field_escape => (
        "package main; import \"fmt\"; import \"html/template\"; import \"bytes\"; func main() { type Inner struct { V string }; type Outer struct { Inner Inner }; t := template.Must(template.New(\"p\").Parse(\"{{.Inner.V}}\")); var b bytes.Buffer; t.Execute(&b, Outer{Inner: Inner{V: \"<z>\"}}); fmt.Println(b.String()) }",
        vec!["&lt;z&gt;"]
    ),
}

go_compile_cases! {
    html_template_must_parse_valid => "package main; import \"html/template\"; func main() { _ = template.Must(template.New(\"m\").Parse(\"{{.}}\")) }",
    html_template_must_parse_with_func => "package main; import \"html/template\"; import \"strings\"; func main() { _ = template.Must(template.New(\"m\").Funcs(template.FuncMap{\"U\": strings.ToUpper}).Parse(\"{{U .}}\")) }",
    html_template_must_parse_files => "package main; import \"html/template\"; func main() { _ = template.Must(template.ParseFiles(\"page.html\")) }",
    html_template_must_parse_glob => "package main; import \"html/template\"; func main() { _ = template.Must(template.ParseGlob(\"*.html\")) }",
    html_template_parse_invalid => "package main; import \"html/template\"; func main() { _, err := template.New(\"p\").Parse(\"{{\"); _ = err }",
    html_template_new_empty => "package main; import \"html/template\"; func main() { _ = template.New(\"root\") }",
    html_template_clone => "package main; import \"html/template\"; func main() { t := template.Must(template.New(\"c\").Parse(\"{{.}}\")); _, _ = t.Clone() }",
    html_template_lookup => "package main; import \"html/template\"; func main() { t := template.Must(template.New(\"n\").Parse(\"{{.}}\")); _ = t.Lookup(\"n\") }",
    html_template_defined_templates => "package main; import \"html/template\"; func main() { t := template.Must(template.New(\"d\").Parse(\"{{.}}\")); _ = t.DefinedTemplates() }",
    html_template_option_missingkey => "package main; import \"html/template\"; func main() { _ = template.New(\"o\").Option(\"missingkey=zero\") }",
    html_template_html_escape_writer => "package main; import \"html/template\"; import \"bytes\"; func main() { template.HTMLEscape(bytes.NewBuffer(nil), []byte(\"<p>\")) }",
    html_template_js_escape_writer => "package main; import \"html/template\"; import \"bytes\"; func main() { template.JSEscape(bytes.NewBuffer(nil), []byte(\"fn()\")) }",
    html_template_css_escape_writer => "package main; import \"html/template\"; import \"bytes\"; func main() { template.CSSEscape(bytes.NewBuffer(nil), []byte(\"body{}\")) }",
    html_template_parse_with_actions => "package main; import \"html/template\"; func main() { _, _ = template.New(\"a\").Parse(`{{define \"t\"}}{{.}}{{end}}`) }",
    html_template_execute_to_writer => "package main; import \"html/template\"; import \"bytes\"; func main() { t := template.Must(template.New(\"e\").Parse(\"{{.}}\")); _ = t.Execute(bytes.NewBuffer(nil), \"x\") }",
    html_template_js_type_in_script => "package main; import \"html/template\"; import \"bytes\"; func main() { t := template.Must(template.New(\"s\").Parse(`<script>{{.}}</script>`)); _ = t.Execute(bytes.NewBuffer(nil), template.JS(\"1+1\")) }",
    html_template_url_type_in_href => "package main; import \"html/template\"; import \"bytes\"; func main() { t := template.Must(template.New(\"u\").Parse(`<a href=\"{{.}}\">l</a>`)); _ = t.Execute(bytes.NewBuffer(nil), template.URL(\"/path\")) }",
    html_template_css_type_in_style => "package main; import \"html/template\"; import \"bytes\"; func main() { t := template.Must(template.New(\"c\").Parse(`<style>{{.}}</style>`)); _ = t.Execute(bytes.NewBuffer(nil), template.CSS(\"p{color:red}\")) }",
    html_template_html_type_raw => "package main; import \"html/template\"; import \"bytes\"; func main() { t := template.Must(template.New(\"h\").Parse(`{{.}}`)); _ = t.Execute(bytes.NewBuffer(nil), template.HTML(\"<br>\")) }",
    html_template_attr_type => "package main; import \"html/template\"; import \"bytes\"; func main() { t := template.Must(template.New(\"a\").Parse(`<div class=\"{{.}}\">`)); _ = t.Execute(bytes.NewBuffer(nil), template.HTMLAttr(\"safe\")) }",
    html_template_parse_tree => "package main; import \"html/template\"; import \"text/template/parse\"; func main() { t := template.New(\"pt\"); tree, _ := parse.Parse(\"pt\", \"{{.}}\", \"\", \"\"); _ = t.AddParseTree(\"pt\", tree) }",
    html_escape_writer => "package main; import \"html\"; import \"bytes\"; func main() { _ = html.Escape(bytes.NewBuffer(nil), []byte(\"<a>\")) }",
    html_unescape_string_amp => "package main; import \"html\"; func main() { _ = html.UnescapeString(\"&amp;\") }",
    html_template_must_new_parse => "package main; import \"html/template\"; func main() { _ = template.Must(template.New(\"x\").Parse(\"ok\")) }",
    html_template_associate => "package main; import \"html/template\"; func main() { t1 := template.Must(template.New(\"a\").Parse(\"{{.}}\")); t2 := template.Must(template.New(\"b\").Parse(\"{{.}}\")); _, _ = t1.AddParseTree(\"b\", t2.Tree) }",
    html_template_execute_template_missing => "package main; import \"html/template\"; import \"bytes\"; func main() { t := template.Must(template.New(\"r\").Parse(\"{{.}}\")); _ = t.ExecuteTemplate(bytes.NewBuffer(nil), \"missing\", nil) }",
    html_template_name => "package main; import \"html/template\"; func main() { t := template.New(\"named\"); _ = t.Name() }",
}
