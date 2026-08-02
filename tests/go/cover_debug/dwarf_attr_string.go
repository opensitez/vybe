// vybe-test: go/cover_debug/dwarf_attr_string
// origin: languages/go/tests/go/test_cover_debug.rs
// vybe-test-mode: compile

package main
import "debug/dwarf"
func main() { _ = dwarf.Attr(0).String() }
