// vybe-test: go/cover_debug_formats/gosym_sym_value
// origin: languages/go/tests/go/test_cover_debug_formats.rs
// vybe-test-mode: compile

package main
import "debug/gosym"
func main() { var s gosym.Sym
_ = s.Value }
