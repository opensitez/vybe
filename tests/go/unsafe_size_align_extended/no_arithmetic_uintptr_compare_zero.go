// vybe-test: go/unsafe_size_align_extended/no_arithmetic_uintptr_compare_zero
// origin: languages/go/tests/go/test_unsafe_size_align_extended.rs
// vybe-test-mode: compile

package main
import "unsafe"
func main() { var p *int
_ = uintptr(unsafe.Pointer(p)) == 0 }
