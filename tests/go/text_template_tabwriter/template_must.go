// vybe-test: go/text_template_tabwriter/template_must
// origin: languages/go/tests/go/test_text_template_tabwriter.rs
// vybe-test-mode: compile

package main
import "text/template"
func main() { _ = template.Must(template.New("t").Parse("x")) }
