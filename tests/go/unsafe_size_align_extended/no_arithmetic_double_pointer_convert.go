// vybe-test: go/unsafe_size_align_extended/no_arithmetic_double_pointer_convert
// origin: languages/go/tests/go/test_unsafe_size_align_extended.rs
// vybe-test-mode: compile

package main
import "unsafe"
func main() { var x int
p := &x
_ = unsafe.Pointer(p) }
