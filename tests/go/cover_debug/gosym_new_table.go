// vybe-test: go/cover_debug/gosym_new_table
// origin: languages/go/tests/go/test_cover_debug.rs
// vybe-test-mode: compile

package main
import "debug/gosym"
import "bytes"
func main() { _, _ = gosym.NewTable(nil, bytes.NewReader(nil)) }
