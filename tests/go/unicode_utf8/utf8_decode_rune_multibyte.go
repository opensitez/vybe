// vybe-test: go/unicode_utf8/utf8_decode_rune_multibyte
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

func main() { r, size := utf8.DecodeRune([]byte("世"))
__p(fmt.Sprint(int(r)))
__p(fmt.Sprint(size)) 
__check("19990\n3")
}
