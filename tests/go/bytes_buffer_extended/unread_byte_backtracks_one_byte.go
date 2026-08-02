// vybe-test: go/bytes_buffer_extended/unread_byte_backtracks_one_byte
// origin: languages/go/tests/go/test_bytes_buffer_extended.rs

package main
import "fmt"
import "bytes"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var b bytes.Buffer
b.WriteString("go")
b.ReadByte()
b.UnreadByte()
ch, _ := b.ReadByte()
__check(fmt.Sprint(string(ch)), "g") }
