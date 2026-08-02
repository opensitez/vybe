// vybe-test: go/cover_debug_formats/macho_open
// origin: languages/go/tests/go/test_cover_debug_formats.rs
// vybe-test-mode: compile

package main
import "debug/macho"
func main() { _, _ = macho.Open("/dev/null") }
