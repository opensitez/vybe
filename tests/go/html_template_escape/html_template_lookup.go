// vybe-test: go/html_template_escape/html_template_lookup
// origin: languages/go/tests/go/test_html_template_escape.rs
// vybe-test-mode: compile

package main
import "html/template"
func main() { t := template.Must(template.New("n").Parse("{{.}}"))
_ = t.Lookup("n") }
