// vybe-test: go/bufio_io/reader_buffered_after_prefetch
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

func main() { r := bufio.NewReaderSize(strings.NewReader("abcd"), 8)
r.ReadByte()
__check(fmt.Sprint(r.Buffered()), "3") }
