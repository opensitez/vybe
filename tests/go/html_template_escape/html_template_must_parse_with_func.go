// vybe-test: go/html_template_escape/html_template_must_parse_with_func
// origin: languages/go/tests/go/test_html_template_escape.rs
// vybe-test-mode: compile

package main
import "html/template"
import "strings"
func main() { _ = template.Must(template.New("m").Funcs(template.FuncMap{"U": strings.ToUpper}).Parse("{{U .}}")) }
