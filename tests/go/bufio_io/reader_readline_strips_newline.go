// vybe-test: go/bufio_io/reader_readline_strips_newline
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

func main() { r := bufio.NewReader(strings.NewReader("hello\n"))
line, _, _ := r.ReadLine()
__check(fmt.Sprint(string(line)), "hello") }
