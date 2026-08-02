// vybe-test: go/text_template_tabwriter/template_func_map
// origin: languages/go/tests/go/test_text_template_tabwriter.rs
// vybe-test-mode: compile

package main
import "text/template"
import "strings"
func main() { _, _ = template.New("t").Funcs(template.FuncMap{"U": strings.ToUpper}).Parse("{{U .}}") }
