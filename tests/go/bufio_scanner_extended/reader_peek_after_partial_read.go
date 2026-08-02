// vybe-test: go/bufio_scanner_extended/reader_peek_after_partial_read
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

func main() { r := bufio.NewReader(strings.NewReader("abcd"))
r.ReadByte()
p, _ := r.Peek(2)
__check(fmt.Sprint(string(p)), "bc") }
