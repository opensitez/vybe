// vybe-test: go/encoding_binary/binary_little_endian_put_uint16_byte_layout
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

func main() { buf := make([]byte, 2)
binary.LittleEndian.PutUint16(buf, 0x0102)
__check(fmt.Sprint(int(buf[0])), "2")
__check(fmt.Sprint(int(buf[1])), "1") }
