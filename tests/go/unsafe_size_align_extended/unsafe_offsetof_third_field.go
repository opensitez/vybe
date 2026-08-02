// vybe-test: go/unsafe_size_align_extended/unsafe_offsetof_third_field
// origin: languages/go/tests/go/test_unsafe_size_align_extended.rs
// vybe-test-mode: compile

package main
import "unsafe"
type S struct { a byte
b byte
c int32 }
func main() { _ = unsafe.Offsetof(S{}.c) }
