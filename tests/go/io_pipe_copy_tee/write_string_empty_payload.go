// vybe-test: go/io_pipe_copy_tee/write_string_empty_payload
// origin: languages/go/tests/go/test_io_pipe_copy_tee.rs

package main
import "fmt"
import "bytes"
import "io"
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

func main() { var buf bytes.Buffer
n, _ := io.WriteString(&buf, "")
__p(fmt.Sprint(n) + " " + fmt.Sprint(buf.Len())) 
__check("0 0")
}
