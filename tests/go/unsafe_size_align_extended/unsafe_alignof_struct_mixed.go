// vybe-test: go/unsafe_size_align_extended/unsafe_alignof_struct_mixed
// origin: languages/go/tests/go/test_unsafe_size_align_extended.rs
// vybe-test-mode: compile

package main
import "unsafe"
type S struct { a int16
b int32
c byte }
func main() { _ = unsafe.Alignof(S{}) }
