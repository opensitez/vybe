// vybe-test: go/cover_debug/gosym_sym_kind_string
// origin: languages/go/tests/go/test_cover_debug.rs
// vybe-test-mode: compile

package main
import "debug/gosym"
func main() { _ = gosym.SymKind(0).String() }
