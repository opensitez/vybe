// vybe-test: go/bufio_scanner_extended/reader_read_string_until_tab
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

func main() { r := bufio.NewReader(strings.NewReader("key\tval"))
s, _ := r.ReadString('\t')
__check(fmt.Sprint(s), "key\t") }
