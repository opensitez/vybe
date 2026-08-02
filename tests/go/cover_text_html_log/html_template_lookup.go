// vybe-test: go/cover_text_html_log/html_template_lookup
// origin: languages/go/tests/go/test_cover_text_html_log.rs
// vybe-test-mode: compile

package main
import "html/template"
func main() { t := html/template.Must(html/template.New("root").Parse("{{.}}"))
_ = t.Lookup("root") }
