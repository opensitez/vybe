// vybe-test: go/io_fs_extended/io_read_byte
// origin: languages/go/tests/go/test_io_fs_extended.rs

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

func main() { r := strings.NewReader("A")
b, err := r.ReadByte()
__p(fmt.Sprint(string(b)) + " " + fmt.Sprint(err == nil)) 
__check("A true")
}
