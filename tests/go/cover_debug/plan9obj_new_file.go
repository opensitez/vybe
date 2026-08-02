// vybe-test: go/cover_debug/plan9obj_new_file
// origin: languages/go/tests/go/test_cover_debug.rs
// vybe-test-mode: compile

package main
import "debug/plan9obj"
import "bytes"
func main() { _, _ = plan9obj.NewFile(bytes.NewReader(nil)) }
