// vybe-test: go/lang_builtins_control/unsafe_add_compile
// origin: languages/go/tests/go/test_lang_builtins_control.rs
// vybe-test-mode: compile

package main
import "unsafe"
func main() { var p *int
_ = unsafe.Add(p, 1) }
