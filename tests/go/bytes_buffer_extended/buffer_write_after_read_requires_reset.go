// vybe-test: go/bytes_buffer_extended/buffer_write_after_read_requires_reset
// origin: languages/go/tests/go/test_bytes_buffer_extended.rs
// vybe-test-mode: compile

package main
import "bytes"
func main() { var b bytes.Buffer
b.WriteString("hi")
b.ReadByte()
b.Reset()
_, _ = b.Write([]byte("ok")) }
