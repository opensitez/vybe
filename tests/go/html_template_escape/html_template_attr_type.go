// vybe-test: go/html_template_escape/html_template_attr_type
// origin: languages/go/tests/go/test_html_template_escape.rs
// vybe-test-mode: compile

package main
import "html/template"
import "bytes"
func main() { t := template.Must(template.New("a").Parse(`<div class="{{.}}">`))
_ = t.Execute(bytes.NewBuffer(nil), template.HTMLAttr("safe")) }
