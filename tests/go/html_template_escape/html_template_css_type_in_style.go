// vybe-test: go/html_template_escape/html_template_css_type_in_style
// origin: languages/go/tests/go/test_html_template_escape.rs
// vybe-test-mode: compile

package main
import "html/template"
import "bytes"
func main() { t := template.Must(template.New("c").Parse(`<style>{{.}}</style>`))
_ = t.Execute(bytes.NewBuffer(nil), template.CSS("p{color:red}")) }
