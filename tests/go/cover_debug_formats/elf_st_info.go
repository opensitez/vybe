// vybe-test: go/cover_debug_formats/elf_st_info
// origin: languages/go/tests/go/test_cover_debug_formats.rs
// vybe-test-mode: compile

package main
import "debug/elf"
func main() { _ = elf.ST_INFO(elf.STB_LOCAL, elf.STT_OBJECT) }
