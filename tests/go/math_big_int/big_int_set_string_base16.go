// vybe-test: go/math_big_int/big_int_set_string_base16
// origin: languages/go/tests/go/test_math_big_int.rs

package main
import "fmt"
import "math/big"
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

func main() { z := new(big.Int)
z.SetString("ff", 16)
__p(fmt.Sprint(z.String())) 
__check("255")
}
