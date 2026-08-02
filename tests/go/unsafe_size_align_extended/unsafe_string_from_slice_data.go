// vybe-test: go/unsafe_size_align_extended/unsafe_string_from_slice_data
// origin: languages/go/tests/go/test_unsafe_size_align_extended.rs
// vybe-test-mode: compile

package main
import "unsafe"
func main() { b := []byte("x")
_ = unsafe.String(unsafe.SliceData(b), len(b)) }
