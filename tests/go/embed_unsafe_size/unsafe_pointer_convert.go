// vybe-test: go/embed_unsafe_size/unsafe_pointer_convert
// origin: languages/go/tests/go/test_embed_unsafe_size.rs
// vybe-test-mode: compile

package main
import "unsafe"
func main() { var x int
p := unsafe.Pointer(&x)
_ = p }
