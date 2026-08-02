// vybe-test: go/cover_text_html_log/html_template_must
// origin: languages/go/tests/go/test_cover_text_html_log.rs
// vybe-test-mode: compile

package main
import "html/template"
func main() { _ = html/template.Must(html/template.New("m").Parse("{{.}}")) }
