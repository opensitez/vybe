// vybe-test: go/io_pipe_copy_tee/write_string_empty_payload
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
n, _ := io.WriteString(&buf, "")
__check(fmt.Sprint(n) + " " + fmt.Sprint(buf.Len()), "0 0") }
