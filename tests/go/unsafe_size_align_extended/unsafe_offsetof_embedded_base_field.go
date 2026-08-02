// vybe-test: go/unsafe_size_align_extended/unsafe_offsetof_embedded_base_field
// origin: languages/go/tests/go/test_unsafe_size_align_extended.rs
// vybe-test-mode: compile

package main
import "unsafe"
type Base struct { id int }
type Wrap struct { Base
extra byte }
func main() { _ = unsafe.Offsetof(Wrap{}.extra) }
