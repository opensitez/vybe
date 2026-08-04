// vybe-test: go/slices_maps/slice_passed_to_func
// origin: languages/go/tests/go/test_slices_maps.rs

package main
import "fmt"
func sumSlice(s []int) int { t := 0
for _, v := range s { t = t + v }
return t } var __buf string

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

func main() { __p(fmt.Sprint(sumSlice([]int{1, 2, 3})))
__check("6")
}
