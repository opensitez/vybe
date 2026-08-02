// vybe-test: go/strings_builder/reader_seek_end
// origin: languages/go/tests/go/test_strings_builder.rs
// vybe-test-mode: compile

package main
import "strings"
func main() { r := strings.NewReader("go")
_, _ = r.Seek(0, 2) }
