// vybe-test: go/cover_debug/elf_new_file
// origin: languages/go/tests/go/test_cover_debug.rs
// vybe-test-mode: compile

package main
import "debug/elf"
import "bytes"
func main() { _, _ = elf.NewFile(bytes.NewReader(nil)) }
