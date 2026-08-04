// vybe-test: go/unicode_utf8/utf8_encode_rune_three_byte
// origin: languages/go/tests/go/test_unicode_utf8.rs

package main
import "fmt"
import "unicode/utf8"
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

func main() { buf := make([]byte, 4)
n := utf8.EncodeRune(buf, '世')
__p(fmt.Sprint(n))
__p(fmt.Sprint(int(buf[0])))
__p(fmt.Sprint(int(buf[1])))
__p(fmt.Sprint(int(buf[2]))) 
__check("3\n228\n184\n150")
}
