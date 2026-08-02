// vybe-test: go/bufio_scanner_extended/reader_peek_zero_length
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
p, _ := r.Peek(0)
__check(fmt.Sprint(len(p)), "0") }
