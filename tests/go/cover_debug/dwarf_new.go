// vybe-test: go/cover_debug/dwarf_new
// origin: languages/go/tests/go/test_cover_debug.rs
// vybe-test-mode: compile

package main
import "debug/dwarf"
import "bytes"
func main() { _, _ = dwarf.New(bytes.NewReader(nil)) }
