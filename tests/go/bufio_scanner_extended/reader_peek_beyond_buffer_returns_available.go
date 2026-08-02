// vybe-test: go/bufio_scanner_extended/reader_peek_beyond_buffer_returns_available
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
p, _ := r.Peek(10)
__check(fmt.Sprint(len(p)) + " " + fmt.Sprint(string(p)), "2 ab") }
