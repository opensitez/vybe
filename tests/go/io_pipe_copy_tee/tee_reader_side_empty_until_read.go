// vybe-test: go/io_pipe_copy_tee/tee_reader_side_empty_until_read
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
tr := io.TeeReader(strings.NewReader("x"), &side)
__check(fmt.Sprint(side.Len()), "0")
_, _ = io.ReadAll(tr)
__check(fmt.Sprint(side.String()), "x") }
