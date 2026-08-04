// vybe-test: go/bufio_io/writer_reset_rebinds_output
// origin: languages/go/tests/go/test_bufio_io.rs

package main
import "fmt"
import "bufio"
import "bytes"
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

func main() { var a bytes.Buffer
var b bytes.Buffer
w := bufio.NewWriter(&a)
w.WriteString("old")
w.Reset(&b)
w.WriteString("new")
w.Flush()
__p(fmt.Sprint(a.String()))
__p(fmt.Sprint(b.String())) 
__check("\nnew")
}
