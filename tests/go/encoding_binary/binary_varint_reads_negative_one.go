// vybe-test: go/encoding_binary/binary_varint_reads_negative_one
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

func main() { v, n := binary.Varint([]byte{0x01})
__check(fmt.Sprint(v), "-1")
__check(fmt.Sprint(n), "1") }
