// vybe-test: go/io_pipe_copy_tee/read_all_binary_null_byte
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

func main() { data, _ := io.ReadAll(strings.NewReader("\x00a"))
__check(fmt.Sprint(len(data)) + " " + fmt.Sprint(int(data[0])), "2 0") }
