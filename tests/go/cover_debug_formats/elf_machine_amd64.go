// vybe-test: go/cover_debug_formats/elf_machine_amd64
// origin: languages/go/tests/go/test_cover_debug_formats.rs
// vybe-test-mode: compile

package main
import "debug/elf"
func main() { _ = elf.EM_X86_64 }
