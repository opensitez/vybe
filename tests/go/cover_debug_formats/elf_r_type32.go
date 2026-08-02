// vybe-test: go/cover_debug_formats/elf_r_type32
// origin: languages/go/tests/go/test_cover_debug_formats.rs
// vybe-test-mode: compile

package main
import "debug/elf"
func main() { _ = elf.R_TYPE32(1) }
