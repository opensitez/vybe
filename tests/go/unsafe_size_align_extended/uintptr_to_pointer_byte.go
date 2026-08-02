// vybe-test: go/unsafe_size_align_extended/uintptr_to_pointer_byte
// origin: languages/go/tests/go/test_unsafe_size_align_extended.rs
// vybe-test-mode: compile

package main
import "unsafe"
func main() { var b byte
u := uintptr(unsafe.Pointer(&b))
_ = (*byte)(unsafe.Pointer(u)) }
