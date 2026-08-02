// vybe-test: go/cover_debug_formats/pe_machine_amd64
// origin: languages/go/tests/go/test_cover_debug_formats.rs
// vybe-test-mode: compile

package main
import "debug/pe"
func main() { _ = pe.IMAGE_FILE_MACHINE_AMD64 }
