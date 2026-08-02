// vybe-test: go/cover_debug_formats/dwarf_attr_sibling
// origin: languages/go/tests/go/test_cover_debug_formats.rs
// vybe-test-mode: compile

package main
import "debug/dwarf"
func main() { _ = dwarf.AttrSibling }
