// vybe-test: go/bufio_scanner_extended/reader_unread_byte_then_read_rune
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

func main() { r := bufio.NewReader(strings.NewReader("go"))
b, _ := r.ReadByte()
r.UnreadByte()
ch, _, _ := r.ReadRune()
__check(fmt.Sprint(string(b)) + " " + fmt.Sprint(string(ch)), "g g") }
