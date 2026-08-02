// vybe-test: go/cover_debug_formats/gosym_line_table_line_to_pc
// origin: languages/go/tests/go/test_cover_debug_formats.rs
// vybe-test-mode: compile

package main
import "debug/gosym"
func main() { t := gosym.NewLineTable(nil, 0)
_ = t.LineToPC(1, 0) }
