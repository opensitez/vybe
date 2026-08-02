// vybe-test: go/bytes_buffer_extended/buffer_read_byte_empty_eof
// origin: languages/go/tests/go/test_bytes_buffer_extended.rs
// vybe-test-mode: compile

package main
import "bytes"
func main() { var b bytes.Buffer
_, _ = b.ReadByte() }
