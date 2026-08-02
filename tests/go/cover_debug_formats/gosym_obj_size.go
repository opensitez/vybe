// vybe-test: go/cover_debug_formats/gosym_obj_size
// origin: languages/go/tests/go/test_cover_debug_formats.rs
// vybe-test-mode: compile

package main
import "debug/gosym"
func main() { var o gosym.Obj
_ = o.Size }
