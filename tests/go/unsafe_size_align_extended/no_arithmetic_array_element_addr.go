// vybe-test: go/unsafe_size_align_extended/no_arithmetic_array_element_addr
// origin: languages/go/tests/go/test_unsafe_size_align_extended.rs
// vybe-test-mode: compile

package main
import "unsafe"
func main() { a := [3]int{1,2,3}
_ = unsafe.Pointer(&a[2]) }
