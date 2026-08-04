// vybe-test: go/io_pipe_copy_tee/read_full_vs_read_at_least_same_buffer
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

func main() { buf1 := make([]byte, 2)
buf2 := make([]byte, 2)
_, e1 := io.ReadFull(strings.NewReader("xy"), buf1)
_, e2 := io.ReadAtLeast(strings.NewReader("xy"), buf2, 2)
__p(fmt.Sprint(string(buf1)) + " " + fmt.Sprint(string(buf2)) + " " + fmt.Sprint(e1 == nil) + " " + fmt.Sprint(e2 == nil)) 
__check("xy xy true true")
}
