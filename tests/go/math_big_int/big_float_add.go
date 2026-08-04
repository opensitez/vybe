// vybe-test: go/math_big_int/big_float_add
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

func main() { a := big.NewFloat(1.5)
b := big.NewFloat(2.5)
__p(fmt.Sprint(a.Add(a, b).String())) 
__check("4")
}
