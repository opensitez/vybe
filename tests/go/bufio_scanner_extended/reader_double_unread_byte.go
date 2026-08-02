// vybe-test: go/bufio_scanner_extended/reader_double_unread_byte
// origin: languages/go/tests/go/test_bufio_scanner_extended.rs

package main
import "fmt"
import "bufio"
import "strings"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { r := bufio.NewReader(strings.NewReader("ab"))
b1, _ := r.ReadByte()
r.UnreadByte()
b2, _ := r.ReadByte()
__check(fmt.Sprint(string(b1)) + " " + fmt.Sprint(string(b2)), "a a") }
