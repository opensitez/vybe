// vybe-test: go/math_big_int/big_float_sub
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

func main() { a := big.NewFloat(5.0)
b := big.NewFloat(2.0)
__p(fmt.Sprint(a.Sub(a, b).String())) 
__check("3")
}
