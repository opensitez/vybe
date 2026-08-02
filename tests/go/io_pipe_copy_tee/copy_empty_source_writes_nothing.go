// vybe-test: go/io_pipe_copy_tee/copy_empty_source_writes_nothing
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
n, err := io.Copy(&dst, strings.NewReader(""))
__check(fmt.Sprint(n) + " " + fmt.Sprint(err == nil) + " " + fmt.Sprint(dst.Len()), "0 true 0") }
