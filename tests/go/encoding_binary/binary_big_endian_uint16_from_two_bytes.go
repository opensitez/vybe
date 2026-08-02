// vybe-test: go/encoding_binary/binary_big_endian_uint16_from_two_bytes
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

func main() { __check(fmt.Sprint(binary.BigEndian.Uint16([]byte{0x01, 0x02})), "258") }
