// vybe-test: go/cover_debug/elf_machine_string
// origin: languages/go/tests/go/test_cover_debug.rs
// vybe-test-mode: compile

package main
import "debug/elf"
func main() { _ = elf.Machine(0).String() }
