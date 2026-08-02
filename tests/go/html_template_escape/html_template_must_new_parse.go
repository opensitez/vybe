// vybe-test: go/html_template_escape/html_template_must_new_parse
// origin: languages/go/tests/go/test_html_template_escape.rs
// vybe-test-mode: compile

package main
import "html/template"
func main() { _ = template.Must(template.New("x").Parse("ok")) }
