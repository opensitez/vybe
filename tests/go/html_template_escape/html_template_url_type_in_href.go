// vybe-test: go/html_template_escape/html_template_url_type_in_href
// origin: languages/go/tests/go/test_html_template_escape.rs
// vybe-test-mode: compile

package main
import "html/template"
import "bytes"
func main() { t := template.Must(template.New("u").Parse(`<a href="{{.}}">l</a>`))
_ = t.Execute(bytes.NewBuffer(nil), template.URL("/path")) }
