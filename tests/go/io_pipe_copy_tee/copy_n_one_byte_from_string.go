// vybe-test: go/io_pipe_copy_tee/copy_n_one_byte_from_string
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
n, _ := io.CopyN(&dst, strings.NewReader("go"), 1)
__check(fmt.Sprint(n) + " " + fmt.Sprint(dst.String()), "1 g") }
