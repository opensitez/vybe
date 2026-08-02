// vybe-test: go/unsafe_size_align_extended/unsafe_alignof_complex64
// origin: languages/go/tests/go/test_unsafe_size_align_extended.rs
// vybe-test-mode: compile

package main
import "unsafe"
func main() { _ = unsafe.Alignof(complex64(0)) }
