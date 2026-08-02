// vybe-test: go/unsafe_size_align_extended/no_arithmetic_func_pointer
// origin: languages/go/tests/go/test_unsafe_size_align_extended.rs
// vybe-test-mode: compile

package main
import "unsafe"
func main() { f := func() {}
_ = unsafe.Pointer(&f) }
