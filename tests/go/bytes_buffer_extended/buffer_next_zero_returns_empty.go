// vybe-test: go/bytes_buffer_extended/buffer_next_zero_returns_empty
// origin: languages/go/tests/go/test_bytes_buffer_extended.rs
// vybe-test-mode: compile

package main
import "bytes"
func main() { var b bytes.Buffer
b.WriteString("x")
_ = b.Next(0) }
