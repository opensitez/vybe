// vybe-test: go/encoding_binary/binary_little_endian_uint32_from_four_bytes
// origin: languages/go/tests/go/test_encoding_binary.rs

package main
import "fmt"
import "encoding/binary"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(binary.LittleEndian.Uint32([]byte{0x2a, 0, 0, 0})), "42") }
