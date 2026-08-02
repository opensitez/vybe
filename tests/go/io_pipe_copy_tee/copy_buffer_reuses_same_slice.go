// vybe-test: go/io_pipe_copy_tee/copy_buffer_reuses_same_slice
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
buf := make([]byte, 3)
_, _ = io.CopyBuffer(&dst, strings.NewReader("xyz"), buf)
__check(fmt.Sprint(len(buf)) + " " + fmt.Sprint(dst.String()), "3 xyz") }
