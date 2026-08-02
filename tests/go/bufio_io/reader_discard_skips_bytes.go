// vybe-test: go/bufio_io/reader_discard_skips_bytes
// origin: languages/go/tests/go/test_bufio_io.rs

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

func main() { r := bufio.NewReader(strings.NewReader("skipme"))
r.Discard(4)
b, _ := r.ReadByte()
__check(fmt.Sprint(string(b)), "m") }
