// vybe-test: go/io_pipe_copy_tee/write_string_unicode_via_interface
// origin: languages/go/tests/go/test_io_pipe_copy_tee.rs

package main
import "fmt"
import "bytes"
import "io"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var buf bytes.Buffer
var w io.Writer = &buf
_, _ = io.WriteString(w, "日")
__check(fmt.Sprint(buf.String()), "日") }
