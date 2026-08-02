// vybe-test: go/bufio_scanner_extended/reader_read_bytes_until_comma
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

func main() { r := bufio.NewReader(strings.NewReader("one,two"))
b, _ := r.ReadBytes(',')
__check(fmt.Sprint(string(b)), "one,") }
