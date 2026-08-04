// vybe-test: go/bufio_io/writer_buffered_before_flush
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

func main() { var buf bytes.Buffer
w := bufio.NewWriter(&buf)
w.WriteString("ab")
__p(fmt.Sprint(w.Buffered()))
w.Flush()
__p(fmt.Sprint(buf.String())) 
__check("2\nab")
}
