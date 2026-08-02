// vybe-test: go/unsafe_size_align_extended/uintptr_from_slice_data
// origin: languages/go/tests/go/test_unsafe_size_align_extended.rs
// vybe-test-mode: compile

package main
import "unsafe"
func main() { sl := []byte{1}
_ = uintptr(unsafe.Pointer(unsafe.SliceData(sl))) }
