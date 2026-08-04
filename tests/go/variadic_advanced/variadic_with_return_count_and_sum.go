// vybe-test: go/variadic_advanced/variadic_with_return_count_and_sum
// origin: languages/go/tests/go/test_variadic_advanced.rs

package main
import "fmt"
func stats(nums ...int) (int, int) { t := 0
for _, n := range nums { t += n }
return len(nums), t }
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

func main() { c, s := stats(2, 3, 4)
__p(fmt.Sprint(c))
__p(fmt.Sprint(s)) 
__check("3\n9")
}
