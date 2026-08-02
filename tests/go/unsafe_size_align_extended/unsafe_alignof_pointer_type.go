// vybe-test: go/unsafe_size_align_extended/unsafe_alignof_pointer_type
// origin: languages/go/tests/go/test_unsafe_size_align_extended.rs
// vybe-test-mode: compile

package main
import "unsafe"
func main() { var p *byte
_ = unsafe.Alignof(p) }
