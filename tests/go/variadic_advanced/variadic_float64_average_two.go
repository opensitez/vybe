// vybe-test: go/variadic_advanced/variadic_float64_average_two
// origin: languages/go/tests/go/test_variadic_advanced.rs

package main
import "fmt"
func avg(nums ...float64) float64 { if len(nums) == 0 { return 0 }
s := 0.0
for _, n := range nums { s += n }
return s / float64(len(nums)) }
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

func main() { __p(fmt.Sprint(avg(2.0, 4.0))) 
__check("3")
}
