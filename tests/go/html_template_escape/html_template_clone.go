// vybe-test: go/html_template_escape/html_template_clone
// origin: languages/go/tests/go/test_html_template_escape.rs
// vybe-test-mode: compile

package main
import "html/template"
func main() { t := template.Must(template.New("c").Parse("{{.}}"))
_, _ = t.Clone() }
