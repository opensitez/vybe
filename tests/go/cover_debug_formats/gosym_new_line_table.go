// vybe-test: go/cover_debug_formats/gosym_new_line_table
// origin: languages/go/tests/go/test_cover_debug_formats.rs
// vybe-test-mode: compile

package main
import "debug/gosym"
func main() { _ = gosym.NewLineTable(nil, 0) }
