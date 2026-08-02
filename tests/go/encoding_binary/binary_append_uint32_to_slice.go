// vybe-test: go/encoding_binary/binary_append_uint32_to_slice
// origin: languages/go/tests/go/test_encoding_binary.rs
// vybe-test-mode: compile

package main
import "encoding/binary"
func main() { b := make([]byte, 0, 8)
_ = binary.LittleEndian.AppendUint32(b, 42) }
