// vybe-test: go/io_pipe_copy_tee/write_string_via_writer_interface
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
n, _ := io.WriteString(w, "iface")
__check(fmt.Sprint(n) + " " + fmt.Sprint(buf.String()), "5 iface") }
