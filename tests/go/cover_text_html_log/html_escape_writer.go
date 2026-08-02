// vybe-test: go/cover_text_html_log/html_escape_writer
// origin: languages/go/tests/go/test_cover_text_html_log.rs
// vybe-test-mode: compile

package main
import "html"
import "bytes"
func main() { _ = html.Escape(bytes.NewBuffer(nil), []byte("<a>")) }
