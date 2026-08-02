// vybe-test: go/text_template_tabwriter/template_execute
// origin: languages/go/tests/go/test_text_template_tabwriter.rs
// vybe-test-mode: compile

package main
import "text/template"
import "bytes"
func main() { t, _ := template.New("t").Parse("{{.}}")
var b bytes.Buffer
_ = t.Execute(&b, "hi") }
