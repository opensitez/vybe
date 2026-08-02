// vybe-test: go/cover_text_html_log/html_template_escape_writer
// origin: languages/go/tests/go/test_cover_text_html_log.rs
// vybe-test-mode: compile

package main
import "html/template"
import "bytes"
func main() { html/template.HTMLEscape(bytes.NewBuffer(nil), []byte("<p>")) }
