// vybe-test: go/unsafe_size_align_extended/uintptr_to_pointer_int
// origin: languages/go/tests/go/test_unsafe_size_align_extended.rs
// vybe-test-mode: compile

package main
import "unsafe"
func main() { var x int
u := uintptr(unsafe.Pointer(&x))
_ = (*int)(unsafe.Pointer(u)) }
