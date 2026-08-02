// vybe-test: go/encoding_binary/binary_little_endian_put_int16_negative_one
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
binary.LittleEndian.PutInt16(buf, -1)
__check(fmt.Sprint(int(buf[0])), "255")
__check(fmt.Sprint(int(buf[1])), "255") }
