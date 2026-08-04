// vybe-test: go/unicode_utf16_norm/utf8_decode_last_rune_in_string
// origin: languages/go/tests/go/test_unicode_utf16_norm.rs

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

func main() { s := "ab"
r, size := utf8.DecodeLastRuneInString(s)
__p(fmt.Sprint(int(r)))
__p(fmt.Sprint(size)) 
__check("98\n1")
}
