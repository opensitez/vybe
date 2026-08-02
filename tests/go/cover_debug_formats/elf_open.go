// vybe-test: go/cover_debug_formats/elf_open
// origin: languages/go/tests/go/test_cover_debug_formats.rs
// vybe-test-mode: compile

package main
import "debug/elf"
func main() { _, _ = elf.Open("/dev/null") }
