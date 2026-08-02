// vybe-test: go/encoding_binary/binary_little_endian_put_float64
// origin: languages/go/tests/go/test_encoding_binary.rs
// vybe-test-mode: compile

package main
import "encoding/binary"
import "math"
func main() { buf := make([]byte, 8)
binary.LittleEndian.PutUint64(buf, math.Float64bits(2.5)) }
