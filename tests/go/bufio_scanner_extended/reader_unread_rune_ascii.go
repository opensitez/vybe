// vybe-test: go/bufio_scanner_extended/reader_unread_rune_ascii
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

func main() { r := bufio.NewReader(strings.NewReader("z"))
ch, _, _ := r.ReadRune()
r.UnreadRune()
ch2, _, _ := r.ReadRune()
__check(fmt.Sprint(string(ch)) + " " + fmt.Sprint(string(ch2)), "z z") }
