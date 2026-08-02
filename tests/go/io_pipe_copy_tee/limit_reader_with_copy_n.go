// vybe-test: go/io_pipe_copy_tee/limit_reader_with_copy_n
// origin: languages/go/tests/go/test_io_pipe_copy_tee.rs

package main
import "fmt"
import "io"
import "bytes"
import "strings"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { lr := io.LimitReader(strings.NewReader("abcdef"), 3)
var dst bytes.Buffer
n, _ := io.CopyN(&dst, lr, 2)
__check(fmt.Sprint(n) + " " + fmt.Sprint(dst.String()), "2 ab") }
