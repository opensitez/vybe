// vybe-test: go/cover_debug/macho_new_file
// origin: languages/go/tests/go/test_cover_debug.rs
// vybe-test-mode: compile

package main
import "debug/macho"
import "bytes"
func main() { _, _ = macho.NewFile(bytes.NewReader(nil)) }
