// vybe-test: go/cover_text_html_log/quotedprintable_reader_read
// origin: languages/go/tests/go/test_cover_text_html_log.rs
// vybe-test-mode: compile

package main
import "mime/quotedprintable"
import "strings"
func main() { r := quotedprintable.NewReader(strings.NewReader("=41"))
buf := make([]byte, 4)
_, _ = r.Read(buf) }
