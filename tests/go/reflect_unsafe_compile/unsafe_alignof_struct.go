// vybe-test: go/reflect_unsafe_compile/unsafe_alignof_struct
// origin: languages/go/tests/go/test_reflect_unsafe_compile.rs
// vybe-test-mode: compile

package main
import "unsafe"
type S struct { a int8
b int32 }
func main() { _ = unsafe.Alignof(S{}) }
