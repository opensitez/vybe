// vybe-test: go/io_pipe_copy_tee/read_at_least_short_source_errors
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

func main() { buf := make([]byte, 5)
n, err := io.ReadAtLeast(strings.NewReader("ab"), buf, 4)
__check(fmt.Sprint(n) + " " + fmt.Sprint(err != nil), "2 true") }
