// vybe-test: go/text_template_tabwriter/template_parse_glob
// origin: languages/go/tests/go/test_text_template_tabwriter.rs
// vybe-test-mode: compile

package main
import "text/template"
func main() { _, _ = template.ParseGlob("*.tmpl") }
