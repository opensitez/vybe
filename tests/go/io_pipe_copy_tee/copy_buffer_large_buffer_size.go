// vybe-test: go/io_pipe_copy_tee/copy_buffer_large_buffer_size
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
buf := make([]byte, 64)
n, _ := io.CopyBuffer(&dst, strings.NewReader("tiny"), buf)
__check(fmt.Sprint(n) + " " + fmt.Sprint(dst.String()), "4 tiny") }
