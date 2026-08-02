// vybe-test: go/text_template_tabwriter/template_define
// origin: languages/go/tests/go/test_text_template_tabwriter.rs
// vybe-test-mode: compile

package main
import "text/template"
func main() { t := template.New("root")
_ = t.New("child") }
