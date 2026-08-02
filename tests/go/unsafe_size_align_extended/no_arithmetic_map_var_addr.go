// vybe-test: go/unsafe_size_align_extended/no_arithmetic_map_var_addr
// origin: languages/go/tests/go/test_unsafe_size_align_extended.rs
// vybe-test-mode: compile

package main
import "unsafe"
func main() { m := map[string]int{"a":1}
_ = unsafe.Pointer(&m) }
