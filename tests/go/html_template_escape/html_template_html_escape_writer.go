// vybe-test: go/html_template_escape/html_template_html_escape_writer
// origin: languages/go/tests/go/test_html_template_escape.rs
// vybe-test-mode: compile

package main
import "html/template"
import "bytes"
func main() { template.HTMLEscape(bytes.NewBuffer(nil), []byte("<p>")) }
