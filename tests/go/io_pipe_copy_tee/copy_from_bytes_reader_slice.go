// vybe-test: go/io_pipe_copy_tee/copy_from_bytes_reader_slice
// origin: languages/go/tests/go/test_io_pipe_copy_tee.rs

package main
import "fmt"
import "io"
import "bytes"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var dst bytes.Buffer
_, _ = io.Copy(&dst, bytes.NewReader([]byte("buf")))
__check(fmt.Sprint(dst.String()), "buf") }
