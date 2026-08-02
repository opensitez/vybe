// vybe-test: go/cover_debug/pe_machine_string
// origin: languages/go/tests/go/test_cover_debug.rs
// vybe-test-mode: compile

package main
import "debug/pe"
func main() { _ = pe.Machine(0).String() }
