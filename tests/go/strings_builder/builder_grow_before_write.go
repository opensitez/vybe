// vybe-test: go/strings_builder/builder_grow_before_write
// origin: languages/go/tests/go/test_strings_builder.rs
// vybe-test-mode: compile

package main
import "strings"
func main() { var b strings.Builder
b.Grow(8)
b.WriteString("grow") }
