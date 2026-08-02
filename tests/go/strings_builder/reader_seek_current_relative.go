// vybe-test: go/strings_builder/reader_seek_current_relative
// origin: languages/go/tests/go/test_strings_builder.rs
// vybe-test-mode: compile

package main
import "strings"
func main() { r := strings.NewReader("go")
_, _ = r.ReadByte()
_, _ = r.Seek(1, 1) }
