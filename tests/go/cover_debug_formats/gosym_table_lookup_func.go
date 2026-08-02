// vybe-test: go/cover_debug_formats/gosym_table_lookup_func
// origin: languages/go/tests/go/test_cover_debug_formats.rs
// vybe-test-mode: compile

package main
import "debug/gosym"
func main() { t, _ := gosym.NewTable(nil, nil)
if t != nil { _ = t.LookupFunc("main.main") } }
