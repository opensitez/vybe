// vybe-test: go/lang_declarations_types/uintptr_convert_compile
// origin: languages/go/tests/go/test_lang_declarations_types.rs
// vybe-test-mode: compile

package main
import "unsafe"
func main() { var p *int
_ = uintptr(unsafe.Pointer(p)) }
