// vybe-test: go/variadic_advanced/forward_variadic_to_peer_direct
// origin: languages/go/tests/go/test_variadic_advanced.rs

package main
import "fmt"
func sink(nums ...int) int { t := 0
for _, n := range nums { t += n }
return t }
func relay(nums ...int) int { return sink(nums...) }
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

func main() { __p(fmt.Sprint(relay(1, 2, 3))) 
__check("6")
}
