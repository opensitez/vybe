// vybe-test: go/bufio_io/scanner_bytes_matches_text
// origin: languages/go/tests/go/test_bufio_io.rs

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

func main() { sc := bufio.NewScanner(strings.NewReader("hi"))
sc.Scan()
__p(fmt.Sprint(string(sc.Bytes()) == sc.Text())) 
__check("true")
}
