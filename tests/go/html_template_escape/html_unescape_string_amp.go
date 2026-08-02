// vybe-test: go/html_template_escape/html_unescape_string_amp
// origin: languages/go/tests/go/test_html_template_escape.rs
// vybe-test-mode: compile

package main
import "html"
func main() { _ = html.UnescapeString("&amp;") }
