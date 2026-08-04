// vybe-test: go/variadic_advanced/mixed_two_fixed_int_variadic_sum
// origin: languages/go/tests/go/test_variadic_advanced.rs

package main
import "fmt"
func offset(base int, step int, vals ...int) int { t := base
for _, v := range vals { t += v + step }
return t }
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

func main() { __p(fmt.Sprint(offset(10, 1, 2, 3))) 
__check("18")
}
