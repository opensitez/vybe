// vybe-test: go/cover_text_html_log/quotedprintable_writer_close
// origin: languages/go/tests/go/test_cover_text_html_log.rs
// vybe-test-mode: compile

package main
import "mime/quotedprintable"
import "bytes"
func main() { w := quotedprintable.NewWriter(bytes.NewBuffer(nil))
_ = w.Close() }
