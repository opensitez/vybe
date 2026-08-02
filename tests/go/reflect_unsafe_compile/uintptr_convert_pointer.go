// vybe-test: go/reflect_unsafe_compile/uintptr_convert_pointer
// origin: languages/go/tests/go/test_reflect_unsafe_compile.rs
// vybe-test-mode: compile

package main
import "unsafe"
func main() { var x int
p := &x
_ = uintptr(unsafe.Pointer(p)) }
