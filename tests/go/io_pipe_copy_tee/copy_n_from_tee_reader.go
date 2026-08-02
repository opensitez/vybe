// vybe-test: go/io_pipe_copy_tee/copy_n_from_tee_reader
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

func main() { var side bytes.Buffer
tr := io.TeeReader(strings.NewReader("go"), &side)
var dst bytes.Buffer
n, _ := io.CopyN(&dst, tr, 1)
__check(fmt.Sprint(n) + " " + fmt.Sprint(dst.String()) + " " + fmt.Sprint(side.String()), "1 g g") }
