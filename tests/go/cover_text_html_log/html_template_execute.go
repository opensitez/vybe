// vybe-test: go/cover_text_html_log/html_template_execute
// origin: languages/go/tests/go/test_cover_text_html_log.rs
// vybe-test-mode: compile

package main
import "html/template"
import "bytes"
func main() { t := html/template.Must(html/template.New("e").Parse("{{.}}"))
_ = t.Execute(bytes.NewBuffer(nil), "hi") }
