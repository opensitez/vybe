// vybe-test: go/lang_builtins_control/unsafe_slice_compile
// origin: languages/go/tests/go/test_lang_builtins_control.rs
// vybe-test-mode: compile

package main
import "unsafe"
func main() { _ = unsafe.Slice((*int)(nil), 0) }
