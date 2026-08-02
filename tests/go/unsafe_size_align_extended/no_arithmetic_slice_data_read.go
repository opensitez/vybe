// vybe-test: go/unsafe_size_align_extended/no_arithmetic_slice_data_read
// origin: languages/go/tests/go/test_unsafe_size_align_extended.rs
// vybe-test-mode: compile

package main
import "unsafe"
func main() { b := []byte{10}
ptr := unsafe.SliceData(b)
_ = *ptr }
