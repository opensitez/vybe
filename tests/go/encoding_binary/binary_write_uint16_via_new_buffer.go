// vybe-test: go/encoding_binary/binary_write_uint16_via_new_buffer
// origin: languages/go/tests/go/test_encoding_binary.rs
// vybe-test-mode: compile

package main
import "bytes"
import "encoding/binary"
func main() { buf := bytes.NewBuffer(nil)
_ = binary.Write(buf, binary.BigEndian, uint16(0x0102)) }
