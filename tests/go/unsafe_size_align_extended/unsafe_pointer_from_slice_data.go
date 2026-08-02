// vybe-test: go/unsafe_size_align_extended/unsafe_pointer_from_slice_data
// origin: languages/go/tests/go/test_unsafe_size_align_extended.rs
// vybe-test-mode: compile

package main
import "unsafe"
func main() { sl := []int{1,2}
_ = unsafe.Pointer(unsafe.SliceData(sl)) }
