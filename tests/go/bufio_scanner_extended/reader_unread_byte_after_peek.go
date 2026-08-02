// vybe-test: go/bufio_scanner_extended/reader_unread_byte_after_peek
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

func main() { r := bufio.NewReader(strings.NewReader("xy"))
p, _ := r.Peek(1)
b, _ := r.ReadByte()
__check(fmt.Sprint(string(p)) + " " + fmt.Sprint(string(b)), "x x") }
