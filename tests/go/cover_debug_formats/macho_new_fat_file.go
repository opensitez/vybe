// vybe-test: go/cover_debug_formats/macho_new_fat_file
// origin: languages/go/tests/go/test_cover_debug_formats.rs
// vybe-test-mode: compile

package main
import "debug/macho"
import "bytes"
func main() { _, _ = macho.NewFatFile(bytes.NewReader(nil)) }
