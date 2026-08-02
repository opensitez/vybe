// vybe-test: go/html_template_escape/html_template_parse_with_actions
// origin: languages/go/tests/go/test_html_template_escape.rs
// vybe-test-mode: compile

package main
import "html/template"
func main() { _, _ = template.New("a").Parse(`{{define "t"}}{{.}}{{end}}`) }
