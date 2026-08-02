// vybe-test: go/cover_debug_formats/gosym_func_entry
// origin: languages/go/tests/go/test_cover_debug_formats.rs
// vybe-test-mode: compile

package main
import "debug/gosym"
func main() { var f gosym.Func
_ = f.Entry }
