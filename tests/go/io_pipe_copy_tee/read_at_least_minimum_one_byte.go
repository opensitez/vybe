// vybe-test: go/io_pipe_copy_tee/read_at_least_minimum_one_byte
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

func main() { buf := make([]byte, 4)
n, err := io.ReadAtLeast(strings.NewReader("go"), buf, 1)
__check(fmt.Sprint(n) + " " + fmt.Sprint(string(buf[:n])) + " " + fmt.Sprint(err == nil), "2 go true") }
