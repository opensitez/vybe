// vybe-test: go/bytes_buffer_extended/buffer_available_after_grow
// origin: languages/go/tests/go/test_bytes_buffer_extended.rs
// vybe-test-mode: compile

package main
import "bytes"
func main() { var b bytes.Buffer
b.Grow(16)
_ = b.Available() }
