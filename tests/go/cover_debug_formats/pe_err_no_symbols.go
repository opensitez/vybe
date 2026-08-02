// vybe-test: go/cover_debug_formats/pe_err_no_symbols
// origin: languages/go/tests/go/test_cover_debug_formats.rs
// vybe-test-mode: compile

package main
import "debug/pe"
func main() { _ = pe.ErrNoSymbols }
