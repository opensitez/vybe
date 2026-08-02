// vybe-test: go/encoding_binary/binary_append_uvarint_to_slice
// origin: languages/go/tests/go/test_encoding_binary.rs
// vybe-test-mode: compile

package main
import "encoding/binary"
func main() { b := []byte{}
_ = binary.AppendUvarint(b, 300) }
