// vybe-test: go/encoding_binary/binary_big_endian_int32_from_slice
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

func main() { __check(fmt.Sprint(binary.BigEndian.Int32([]byte{0xff, 0xff, 0xff, 0xff})), "-1") }
