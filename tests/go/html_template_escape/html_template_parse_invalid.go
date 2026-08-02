// vybe-test: go/html_template_escape/html_template_parse_invalid
// origin: languages/go/tests/go/test_html_template_escape.rs
// vybe-test-mode: compile

package main
import "html/template"
func main() { _, err := template.New("p").Parse("{{")
_ = err }
