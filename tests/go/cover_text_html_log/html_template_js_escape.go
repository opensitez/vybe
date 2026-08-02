// vybe-test: go/cover_text_html_log/html_template_js_escape
// origin: languages/go/tests/go/test_cover_text_html_log.rs
// vybe-test-mode: compile

package main
import "html/template"
import "bytes"
func main() { html/template.JSEscape(bytes.NewBuffer(nil), []byte("fn()")) }
