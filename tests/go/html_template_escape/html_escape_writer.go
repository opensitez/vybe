// vybe-test: go/html_template_escape/html_escape_writer
// origin: languages/go/tests/go/test_html_template_escape.rs
// vybe-test-mode: compile

package main
import "html"
import "bytes"
func main() { _ = html.Escape(bytes.NewBuffer(nil), []byte("<a>")) }
