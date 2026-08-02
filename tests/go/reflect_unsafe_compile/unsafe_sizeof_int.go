// vybe-test: go/reflect_unsafe_compile/unsafe_sizeof_int
// origin: languages/go/tests/go/test_reflect_unsafe_compile.rs
// vybe-test-mode: compile

package main
import "unsafe"
func main() { _ = unsafe.Sizeof(int(0)) }
