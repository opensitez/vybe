// vybe-test: go/encoding_binary/binary_read_decodes_struct_fields
// origin: languages/go/tests/go/test_encoding_binary.rs
// vybe-test-mode: compile

package main
import "bytes"
import "encoding/binary"
type Header struct { Magic uint16
Ver uint8 }
func main() { var h Header
_ = binary.Read(bytes.NewReader([]byte{0xbe, 0xef, 0x01}), binary.BigEndian, &h) }
