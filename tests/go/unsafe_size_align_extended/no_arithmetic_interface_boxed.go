// vybe-test: go/unsafe_size_align_extended/no_arithmetic_interface_boxed
// origin: languages/go/tests/go/test_unsafe_size_align_extended.rs
// vybe-test-mode: compile

package main
import "unsafe"
func main() { var v interface{} = 42
_ = unsafe.Pointer(&v) }
