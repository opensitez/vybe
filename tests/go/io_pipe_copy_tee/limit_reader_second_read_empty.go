// vybe-test: go/io_pipe_copy_tee/limit_reader_second_read_empty
// origin: languages/go/tests/go/test_io_pipe_copy_tee.rs

package main
import "fmt"
import "io"
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

func main() { lr := io.LimitReader(strings.NewReader("xy"), 2)
buf := make([]byte, 1)
n1, _ := lr.Read(buf)
n2, _ := lr.Read(buf)
__p(fmt.Sprint(n1) + " " + fmt.Sprint(string(buf[:n1])) + " " + fmt.Sprint(n2)) 
__check("1 y 1")
}
