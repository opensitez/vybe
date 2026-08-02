// vybe-test: go/strings_builder/reader_read_at_offset
// origin: languages/go/tests/go/test_strings_builder.rs
// vybe-test-mode: compile

package main
import "strings"
func main() { r := strings.NewReader("abc")
buf := make([]byte, 1)
_, _ = r.ReadAt(buf, 1) }
