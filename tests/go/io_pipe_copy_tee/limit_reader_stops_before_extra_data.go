// vybe-test: go/io_pipe_copy_tee/limit_reader_stops_before_extra_data
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

func main() { lr := io.LimitReader(strings.NewReader("abcdef"), 4)
data, _ := io.ReadAll(lr)
__check(fmt.Sprint(string(data)), "abcd") }
