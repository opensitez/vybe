// vybe-test: go/encoding_binary/binary_put_uvarint_two_byte_encoding_for_128
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

func main() { buf := make([]byte, binary.MaxVarintLen64)
n := binary.PutUvarint(buf, 128)
__check(fmt.Sprint(n), "2")
__check(fmt.Sprint(int(buf[0])), "128")
__check(fmt.Sprint(int(buf[1])), "1") }
