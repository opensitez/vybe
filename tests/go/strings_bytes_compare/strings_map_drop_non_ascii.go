// vybe-test: go/strings_bytes_compare/strings_map_drop_non_ascii
// origin: languages/go/tests/go/test_strings_bytes_compare.rs
// vybe-test-mode: compile

package main
import "strings"
func main() { _ = strings.Map(func(r rune) rune { if r > 127 { return -1 }; return r }, "a日b") }
