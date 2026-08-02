// vybe-test: go/unsafe_size_align_extended/unsafe_alignof_float64
// origin: languages/go/tests/go/test_unsafe_size_align_extended.rs
// vybe-test-mode: compile

package main
import "unsafe"
func main() { _ = unsafe.Alignof(float64(0)) }
