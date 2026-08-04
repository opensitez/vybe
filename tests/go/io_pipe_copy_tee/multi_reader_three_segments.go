// vybe-test: go/io_pipe_copy_tee/multi_reader_three_segments
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

func main() { mr := io.MultiReader(strings.NewReader("a"), strings.NewReader("b"), strings.NewReader("c"))
data, _ := io.ReadAll(mr)
__p(fmt.Sprint(string(data))) 
__check("abc")
}
