// vybe-test: go/cover_text_html_log/html_escape_string
// origin: languages/go/tests/go/test_cover_text_html_log.rs
// vybe-test-mode: compile

package main
import "html"
func main() { _ = html.EscapeString("<b>bold</b>") }
