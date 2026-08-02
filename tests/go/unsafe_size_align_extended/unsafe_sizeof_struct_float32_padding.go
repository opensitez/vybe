// vybe-test: go/unsafe_size_align_extended/unsafe_sizeof_struct_float32_padding
// origin: languages/go/tests/go/test_unsafe_size_align_extended.rs
// vybe-test-mode: compile

package main
import "unsafe"
type S struct { f float32
b byte }
func main() { _ = unsafe.Sizeof(S{}) }
