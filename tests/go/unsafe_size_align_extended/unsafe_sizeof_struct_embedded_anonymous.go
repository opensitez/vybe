// vybe-test: go/unsafe_size_align_extended/unsafe_sizeof_struct_embedded_anonymous
// origin: languages/go/tests/go/test_unsafe_size_align_extended.rs
// vybe-test-mode: compile

package main
import "unsafe"
type Base struct { id int32 }
type Child struct { Base
name string }
func main() { _ = unsafe.Sizeof(Child{}) }
