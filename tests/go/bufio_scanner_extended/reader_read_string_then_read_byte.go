// vybe-test: go/bufio_scanner_extended/reader_read_string_then_read_byte
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

func main() { r := bufio.NewReader(strings.NewReader("hi\nthere"))
_, _ = r.ReadString('\n')
b, _ := r.ReadByte()
__check(fmt.Sprint(string(b)), "t") }
