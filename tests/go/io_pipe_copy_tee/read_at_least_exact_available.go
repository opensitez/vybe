// vybe-test: go/io_pipe_copy_tee/read_at_least_exact_available
// origin: languages/go/tests/go/test_io_pipe_copy_tee.rs

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
n, err := io.ReadAtLeast(strings.NewReader("abc"), buf, 3)
__check(fmt.Sprint(n) + " " + fmt.Sprint(string(buf)) + " " + fmt.Sprint(err == nil), "3 abc true") }
