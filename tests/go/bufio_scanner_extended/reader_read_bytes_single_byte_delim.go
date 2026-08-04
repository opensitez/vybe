// vybe-test: go/bufio_scanner_extended/reader_read_bytes_single_byte_delim
// origin: languages/go/tests/go/test_bufio_scanner_extended.rs

package main
import "fmt"
import "bufio"
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

func main() { r := bufio.NewReader(strings.NewReader("x:y"))
b, _ := r.ReadBytes(':')
__p(fmt.Sprint(string(b))) 
__check("x:")
}
