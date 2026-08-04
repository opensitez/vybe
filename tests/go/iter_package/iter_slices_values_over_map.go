// vybe-test: go/iter_package/iter_slices_values_over_map
// origin: languages/go/tests/go/test_iter_package.rs

package main
import "fmt"
import "slices"
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

func main() { m := map[int]int{1: 10, 2: 20}
sum := 0
for v := range slices.Values(m) { sum += v }
__p(fmt.Sprint(sum)) 
__check("30")
}
