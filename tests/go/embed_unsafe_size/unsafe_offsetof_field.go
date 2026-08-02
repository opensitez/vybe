// vybe-test: go/embed_unsafe_size/unsafe_offsetof_field
// origin: languages/go/tests/go/test_embed_unsafe_size.rs
// vybe-test-mode: compile

package main
import "unsafe"
type S struct { a int
b int }
func main() { _ = unsafe.Offsetof(S{}.b) }
