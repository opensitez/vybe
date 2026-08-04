// vybe-test: go/strings_bytes_compare/bytes_to_upper_non_ascii_unchanged
// origin: languages/go/tests/go/test_strings_bytes_compare.rs

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

func main() { b := []byte("日")
u := bytes.ToUpper(b)
__p(fmt.Sprint(bytes.Equal(b, u))) 
__check("true")
}
