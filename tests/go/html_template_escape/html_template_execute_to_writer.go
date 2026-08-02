// vybe-test: go/html_template_escape/html_template_execute_to_writer
// origin: languages/go/tests/go/test_html_template_escape.rs
// vybe-test-mode: compile

package main
import "html/template"
import "bytes"
func main() { t := template.Must(template.New("e").Parse("{{.}}"))
_ = t.Execute(bytes.NewBuffer(nil), "x") }
