// vybe-test: go/strings_builder/reader_write_to
// origin: languages/go/tests/go/test_strings_builder.rs
// vybe-test-mode: compile

package main
import "strings"
func main() { r := strings.NewReader("go")
var b strings.Builder
_, _ = r.WriteTo(&b) }
