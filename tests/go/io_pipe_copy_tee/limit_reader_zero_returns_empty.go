// vybe-test: go/io_pipe_copy_tee/limit_reader_zero_returns_empty
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

func main() { data, _ := io.ReadAll(io.LimitReader(strings.NewReader("data"), 0))
__check(fmt.Sprint(len(data)), "0") }
