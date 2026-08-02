// vybe-test: go/reflect_unsafe_compile/unsafe_offsetof_field
// origin: languages/go/tests/go/test_reflect_unsafe_compile.rs
// vybe-test-mode: compile

package main
import "unsafe"
type S struct { a int
b int }
func main() { _ = unsafe.Offsetof(S{}.b) }
