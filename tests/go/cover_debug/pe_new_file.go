// vybe-test: go/cover_debug/pe_new_file
// origin: languages/go/tests/go/test_cover_debug.rs
// vybe-test-mode: compile

package main
import "debug/pe"
import "bytes"
func main() { _, _ = pe.NewFile(bytes.NewReader(nil)) }
