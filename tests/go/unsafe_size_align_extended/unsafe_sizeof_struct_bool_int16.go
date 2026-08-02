// vybe-test: go/unsafe_size_align_extended/unsafe_sizeof_struct_bool_int16
// origin: languages/go/tests/go/test_unsafe_size_align_extended.rs
// vybe-test-mode: compile

package main
import "unsafe"
type S struct { ok bool
n int16 }
func main() { _ = unsafe.Sizeof(S{}) }
