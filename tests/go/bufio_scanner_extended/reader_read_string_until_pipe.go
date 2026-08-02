// vybe-test: go/bufio_scanner_extended/reader_read_string_until_pipe
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

func main() { r := bufio.NewReader(strings.NewReader("a|b"))
s, _ := r.ReadString('|')
__check(fmt.Sprint(s), "a|") }
