// vybe-test: go/embed_unsafe_size/unsafe_sizeof_int
// origin: languages/go/tests/go/test_embed_unsafe_size.rs
// vybe-test-mode: compile

package main
import "unsafe"
func main() { _ = unsafe.Sizeof(int(0)) }
