// vybe-test: go/variadic_advanced/spread_subslice_middle_two
// origin: languages/go/tests/go/test_variadic_advanced.rs

package main
import "fmt"
func sum(nums ...int) int { t := 0
for _, n := range nums { t += n }
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

func main() { all := []int{5, 1, 2, 3, 9}
__p(fmt.Sprint(sum(all[1:3]...))) 
__check("3")
}
