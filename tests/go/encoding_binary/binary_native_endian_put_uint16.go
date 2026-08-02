// vybe-test: go/encoding_binary/binary_native_endian_put_uint16
// origin: languages/go/tests/go/test_encoding_binary.rs
// vybe-test-mode: compile

package main
import "encoding/binary"
func main() { buf := make([]byte, 2)
binary.NativeEndian.PutUint16(buf, 0xabcd) }
