// vybe-test: go/slices_sort_equal_extended/slices_binary_search_func_custom_order
// origin: languages/go/tests/go/test_slices_sort_equal_extended.rs

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

func main() { s := []int{3, 2, 1}
i, ok := slices.BinarySearchFunc(s, 2, func(a, b int) int { if a > b { return -1 }; if a < b { return 1 }; return 0 })
__p(fmt.Sprint(i))
__p(fmt.Sprint(ok)) 
__check("1\ntrue")
}
