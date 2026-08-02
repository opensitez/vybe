// vybe-test: go/html_template_escape/html_template_js_type_in_script
// origin: languages/go/tests/go/test_html_template_escape.rs
// vybe-test-mode: compile

package main
import "html/template"
import "bytes"
func main() { t := template.Must(template.New("s").Parse(`<script>{{.}}</script>`))
_ = t.Execute(bytes.NewBuffer(nil), template.JS("1+1")) }
