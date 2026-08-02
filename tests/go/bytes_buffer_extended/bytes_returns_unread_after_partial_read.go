// vybe-test: go/bytes_buffer_extended/bytes_returns_unread_after_partial_read
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
b.WriteString("abc")
b.ReadByte()
__check(fmt.Sprint(string(b.Bytes())), "bc") }
