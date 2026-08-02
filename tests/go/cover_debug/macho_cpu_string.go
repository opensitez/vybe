// vybe-test: go/cover_debug/macho_cpu_string
// origin: languages/go/tests/go/test_cover_debug.rs
// vybe-test-mode: compile

package main
import "debug/macho"
func main() { _ = macho.Cpu(0).String() }
