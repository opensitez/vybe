// vybe-test: go/encoding_binary/binary_append_uint16_extends_empty_slice
// origin: languages/go/tests/go/test_encoding_binary.rs
// vybe-test-mode: compile

package main
import "encoding/binary"
func main() { b := []byte{}
_ = binary.BigEndian.AppendUint16(b, 0x0102) }
