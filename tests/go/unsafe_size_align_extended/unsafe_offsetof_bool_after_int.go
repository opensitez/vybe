// vybe-test: go/unsafe_size_align_extended/unsafe_offsetof_bool_after_int
// origin: languages/go/tests/go/test_unsafe_size_align_extended.rs
// vybe-test-mode: compile

package main
import "unsafe"
type S struct { n int64
flag bool }
func main() { _ = unsafe.Offsetof(S{}.flag) }
