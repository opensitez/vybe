// vybe-test: go/unsafe_size_align_extended/no_arithmetic_chan_addr
// origin: languages/go/tests/go/test_unsafe_size_align_extended.rs
// vybe-test-mode: compile

package main
import "unsafe"
func main() { ch := make(chan int, 1)
_ = unsafe.Pointer(&ch) }
