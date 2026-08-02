// vybe-test: go/io_pipe_copy_tee/limit_reader_second_read_empty
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

func main() { lr := io.LimitReader(strings.NewReader("xy"), 2)
buf := make([]byte, 1)
n1, _ := lr.Read(buf)
n2, _ := lr.Read(buf)
__check(fmt.Sprint(n1) + " " + fmt.Sprint(string(buf[:n1])) + " " + fmt.Sprint(n2), "1 y 1") }
