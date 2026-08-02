// vybe-test: go/encoding_binary/binary_big_endian_float32_from_bits
// origin: languages/go/tests/go/test_encoding_binary.rs
// vybe-test-mode: compile

package main
import "encoding/binary"
import "math"
func main() { _ = math.Float32frombits(binary.BigEndian.Uint32([]byte{0x3f, 0xc0, 0, 0})) }
