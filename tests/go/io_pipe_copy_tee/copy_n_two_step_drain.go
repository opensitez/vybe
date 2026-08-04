// vybe-test: go/io_pipe_copy_tee/copy_n_two_step_drain
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
src := strings.NewReader("abcd")
n1, _ := io.CopyN(&dst, src, 2)
n2, _ := io.Copy(&dst, src)
__p(fmt.Sprint(n1) + " " + fmt.Sprint(n2) + " " + fmt.Sprint(dst.String())) 
__check("2 2 abcd")
}
