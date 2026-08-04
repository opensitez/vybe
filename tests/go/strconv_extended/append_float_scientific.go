// vybe-test: go/strconv_extended/append_float_scientific
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

func main() { b := strconv.AppendFloat([]byte{}, 100.0, 'e', 0, 64)
__p(fmt.Sprint(len(string(b)) > 0)) 
__check("true")
}
