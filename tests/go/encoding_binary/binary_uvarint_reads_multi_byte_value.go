// vybe-test: go/encoding_binary/binary_uvarint_reads_multi_byte_value
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

func main() { v, n := binary.Uvarint([]byte{0x80, 0x01})
__check(fmt.Sprint(v), "128")
__check(fmt.Sprint(n), "2") }
