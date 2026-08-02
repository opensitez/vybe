// vybe-test: go/unsafe_size_align_extended/unsafe_pointer_convert_to_byte_ptr
// origin: languages/go/tests/go/test_unsafe_size_align_extended.rs
// vybe-test-mode: compile

package main
import "unsafe"
func main() { var x int
_ = (*byte)(unsafe.Pointer(&x)) }
