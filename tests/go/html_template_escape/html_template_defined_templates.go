// vybe-test: go/html_template_escape/html_template_defined_templates
// origin: languages/go/tests/go/test_html_template_escape.rs
// vybe-test-mode: compile

package main
import "html/template"
func main() { t := template.Must(template.New("d").Parse("{{.}}"))
_ = t.DefinedTemplates() }
