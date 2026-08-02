// vybe-test: go/io_pipe_copy_tee/copy_n_two_step_drain
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

func main() { var dst bytes.Buffer
src := strings.NewReader("abcd")
n1, _ := io.CopyN(&dst, src, 2)
n2, _ := io.Copy(&dst, src)
__check(fmt.Sprint(n1) + " " + fmt.Sprint(n2) + " " + fmt.Sprint(dst.String()), "2 2 abcd") }
