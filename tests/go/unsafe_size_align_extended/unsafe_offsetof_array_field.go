// vybe-test: go/unsafe_size_align_extended/unsafe_offsetof_array_field
// origin: languages/go/tests/go/test_unsafe_size_align_extended.rs
// vybe-test-mode: compile

package main
import "unsafe"
type S struct { buf [4]byte
n int }
func main() { _ = unsafe.Offsetof(S{}.n) }
