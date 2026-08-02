// vybe-test: go/encoding_binary/binary_read_full_exact_byte_count
// origin: languages/go/tests/go/test_encoding_binary.rs
// vybe-test-mode: compile

package main
import "bytes"
import "encoding/binary"
func main() { r := bytes.NewReader([]byte{1, 2, 3, 4})
dst := make([]byte, 4)
_, _ = binary.ReadFull(r, dst) }
