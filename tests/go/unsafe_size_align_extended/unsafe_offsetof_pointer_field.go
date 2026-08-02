// vybe-test: go/unsafe_size_align_extended/unsafe_offsetof_pointer_field
// origin: languages/go/tests/go/test_unsafe_size_align_extended.rs
// vybe-test-mode: compile

package main
import "unsafe"
type S struct { p *int
x int }
func main() { _ = unsafe.Offsetof(S{}.x) }
