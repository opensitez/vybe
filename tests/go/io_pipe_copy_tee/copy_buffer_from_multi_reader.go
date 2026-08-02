// vybe-test: go/io_pipe_copy_tee/copy_buffer_from_multi_reader
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
mr := io.MultiReader(strings.NewReader("v"), strings.NewReader("y"))
buf := make([]byte, 1)
_, _ = io.CopyBuffer(&dst, mr, buf)
__check(fmt.Sprint(dst.String()), "vy") }
