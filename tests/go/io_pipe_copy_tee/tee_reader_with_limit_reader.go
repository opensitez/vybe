// vybe-test: go/io_pipe_copy_tee/tee_reader_with_limit_reader
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
src := io.LimitReader(strings.NewReader("long"), 2)
tr := io.TeeReader(src, &side)
data, _ := io.ReadAll(tr)
__check(fmt.Sprint(string(data)) + " " + fmt.Sprint(side.String()), "lo lo") }
