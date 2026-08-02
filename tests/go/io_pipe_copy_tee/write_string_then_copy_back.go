// vybe-test: go/io_pipe_copy_tee/write_string_then_copy_back
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

func main() { var a bytes.Buffer
io.WriteString(&a, "src")
var b bytes.Buffer
_, _ = io.Copy(&b, &a)
__check(fmt.Sprint(b.String()), "src") }
