// vybe-test: go/unsafe_size_align_extended/no_arithmetic_pointer_store_load
// origin: languages/go/tests/go/test_unsafe_size_align_extended.rs
// vybe-test-mode: compile

package main
import "unsafe"
func main() { var x int
p := (*int)(unsafe.Pointer(&x))
*p = 1
_ = *p }
