// vybe-test: go/encoding_binary/binary_big_endian_put_float32
// origin: languages/go/tests/go/test_encoding_binary.rs
// vybe-test-mode: compile

package main
import "encoding/binary"
import "math"
func main() { buf := make([]byte, 4)
binary.BigEndian.PutUint32(buf, math.Float32bits(1.5)) }
