// vybe-test: go/encoding_binary/binary_big_endian_put_uint64_first_and_last_byte
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

func main() { buf := make([]byte, 8)
binary.BigEndian.PutUint64(buf, 0x0102030405060708)
__check(fmt.Sprint(int(buf[0])), "1")
__check(fmt.Sprint(int(buf[7])), "8") }
