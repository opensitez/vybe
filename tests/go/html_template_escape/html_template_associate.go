// vybe-test: go/html_template_escape/html_template_associate
// origin: languages/go/tests/go/test_html_template_escape.rs
// vybe-test-mode: compile

package main
import "html/template"
func main() { t1 := template.Must(template.New("a").Parse("{{.}}"))
t2 := template.Must(template.New("b").Parse("{{.}}"))
_, _ = t1.AddParseTree("b", t2.Tree) }
