// vybe-test: go/cover_text_html_log/html_template_add_parse_tree
// origin: languages/go/tests/go/test_cover_text_html_log.rs
// vybe-test-mode: compile

package main
import "html/template"
import "text/template/parse"
func main() { t := html/template.New("a")
tree, _ := parse.Parse("a", "{{.}}", "", "")
_ = t.AddParseTree("a", tree) }
