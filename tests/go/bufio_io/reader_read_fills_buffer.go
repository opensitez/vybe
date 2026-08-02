// vybe-test: go/bufio_io/reader_read_fills_buffer
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

func main() { r := bufio.NewReader(strings.NewReader("go"))
buf := make([]byte, 10)
n, _ := r.Read(buf)
__check(fmt.Sprint(n), "2")
__check(fmt.Sprint(string(buf[:n])), "go") }
