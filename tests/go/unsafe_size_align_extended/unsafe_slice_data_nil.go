// vybe-test: go/unsafe_size_align_extended/unsafe_slice_data_nil
// origin: languages/go/tests/go/test_unsafe_size_align_extended.rs
// vybe-test-mode: compile

package main
import "unsafe"
func main() { var sl []int
_ = unsafe.SliceData(sl) }
