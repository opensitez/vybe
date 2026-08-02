// vybe-test: go/encoding_binary/binary_read_uint16_from_bytes_reader
// origin: languages/go/tests/go/test_encoding_binary.rs

package main
import "bytes"
import "fmt"
import "encoding/binary"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { r := bytes.NewReader([]byte{0x01, 0x02})
var v uint16
_ = binary.Read(r, binary.BigEndian, &v)
__check(fmt.Sprint(v), "258") }
