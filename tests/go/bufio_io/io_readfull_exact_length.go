// vybe-test: go/bufio_io/io_readfull_exact_length
// origin: languages/go/tests/go/test_bufio_io.rs

package main
import "fmt"
import "io"
import "strings"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { buf := make([]byte, 3)
_, err := io.ReadFull(strings.NewReader("abc"), buf)
__check(fmt.Sprint(string(buf)), "abc")
__check(fmt.Sprint(err == nil), "true") }
