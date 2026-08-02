// vybe-test: go/strings_builder/map_drop_all_runes
// origin: languages/go/tests/go/test_strings_builder.rs
// vybe-test-mode: compile

package main
import "strings"
func main() { _ = strings.Map(func(r rune) rune { return -1 }, "abc") }
