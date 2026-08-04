// vybe-test: go/strconv_extended/parse_int_bitsize_16
// origin: languages/go/tests/go/test_strconv_extended.rs

package main
import "fmt"
import "strconv"
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

func main() { n, _ := strconv.ParseInt("-32768", 10, 16)
__p(fmt.Sprint(n)) 
__check("-32768")
}
