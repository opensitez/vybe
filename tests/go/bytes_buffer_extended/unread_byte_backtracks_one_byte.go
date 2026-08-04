// vybe-test: go/bytes_buffer_extended/unread_byte_backtracks_one_byte
// origin: languages/go/tests/go/test_bytes_buffer_extended.rs

package main
import "fmt"
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

func main() { var b bytes.Buffer
b.WriteString("go")
b.ReadByte()
b.UnreadByte()
ch, _ := b.ReadByte()
__p(fmt.Sprint(string(ch))) 
__check("g")
}
