// vybe-test: go/unsafe_size_align_extended/no_arithmetic_nested_struct_addr
// origin: languages/go/tests/go/test_unsafe_size_align_extended.rs
// vybe-test-mode: compile

package main
import "unsafe"
type S struct { inner [2]byte }
func main() { var s S
_ = unsafe.Pointer(&s.inner[1]) }
