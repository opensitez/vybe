// vybe-test: go/cover_debug_formats/elf_type_exec
// origin: languages/go/tests/go/test_cover_debug_formats.rs
// vybe-test-mode: compile

package main
import "debug/elf"
func main() { _ = elf.ET_EXEC }
