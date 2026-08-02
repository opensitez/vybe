// vybe-test: go/cover_debug_formats/plan9obj_err_no_symbols
// origin: languages/go/tests/go/test_cover_debug_formats.rs
// vybe-test-mode: compile

package main
import "debug/plan9obj"
func main() { _ = plan9obj.ErrNoSymbols }
