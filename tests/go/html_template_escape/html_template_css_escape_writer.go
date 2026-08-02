// vybe-test: go/html_template_escape/html_template_css_escape_writer
// origin: languages/go/tests/go/test_html_template_escape.rs
// vybe-test-mode: compile

package main
import "html/template"
import "bytes"
func main() { template.CSSEscape(bytes.NewBuffer(nil), []byte("body{}")) }
