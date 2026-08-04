// vybe-test: go/io_pipe_copy_tee/copy_buffer_reuses_same_slice
// origin: languages/go/tests/go/test_io_pipe_copy_tee.rs

package main
import "fmt"
import "io"
import "bytes"
import "strings"
var __buf string

// __p appends one line, __pr appends without a newline.
func __p(s string) { __buf = __buf + s + "\n" }

func __pr(s string) { __buf = __buf + s }

// __check ends the program unless the collected output equals want. The final
// Println contributes a trailing newline the expected line vector never
// carried, so both forms are accepted.
func __check(want string) {
	if __buf != want && __buf != want+"\n" {
		fmt.Println("FAIL: want [" + want + "] got [" + __buf + "]")
		panic("assertion failed")
	}
}

func main() { var dst bytes.Buffer
buf := make([]byte, 3)
_, _ = io.CopyBuffer(&dst, strings.NewReader("xyz"), buf)
__p(fmt.Sprint(len(buf)) + " " + fmt.Sprint(dst.String())) 
__check("3 xyz")
}
