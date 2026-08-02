// vybe-test: go/encoding_binary/binary_big_endian_put_uint32_byte_layout
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

func main() { buf := make([]byte, 4)
binary.BigEndian.PutUint32(buf, 0x01020304)
__check(fmt.Sprint(int(buf[0])), "1")
__check(fmt.Sprint(int(buf[3])), "4") }
