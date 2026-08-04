// vybe-test: go/unicode_utf16_norm/utf16_decode_replaces_invalid_surrogate
// origin: languages/go/tests/go/test_unicode_utf16_norm.rs

package main
import "fmt"
import "unicode/utf16"
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

func main() { rs := utf16.Decode([]uint16{0xD800})
__p(fmt.Sprint(int(rs[0]))) 
__check("65533")
}
