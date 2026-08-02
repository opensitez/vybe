// vybe-test: go/html_template_escape/html_template_parse_tree
// origin: languages/go/tests/go/test_html_template_escape.rs
// vybe-test-mode: compile

package main
import "html/template"
import "text/template/parse"
func main() { t := template.New("pt")
tree, _ := parse.Parse("pt", "{{.}}", "", "")
_ = t.AddParseTree("pt", tree) }
