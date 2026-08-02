// vybe-test: go/cover_text_html_log/quotedprintable_new_reader
// origin: languages/go/tests/go/test_cover_text_html_log.rs
// vybe-test-mode: compile

package main
import "mime/quotedprintable"
import "strings"
func main() { _ = quotedprintable.NewReader(strings.NewReader("=41")) }
