// vybe-test: go/unsafe_size_align_extended/unsafe_pointer_from_struct_field
// origin: languages/go/tests/go/test_unsafe_size_align_extended.rs
// vybe-test-mode: compile

package main
import "unsafe"
type S struct { n int }
func main() { var s S
_ = unsafe.Pointer(&s.n) }
