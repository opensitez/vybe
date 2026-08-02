// vybe-test: go/html_template_escape/html_template_html_type_raw
// origin: languages/go/tests/go/test_html_template_escape.rs
// vybe-test-mode: compile

package main
import "html/template"
import "bytes"
func main() { t := template.Must(template.New("h").Parse(`{{.}}`))
_ = t.Execute(bytes.NewBuffer(nil), template.HTML("<br>")) }
