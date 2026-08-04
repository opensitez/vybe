// vybe-test: go/slices_sort_equal_extended/slices_binary_search_found_exact
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

func main() { i, ok := slices.BinarySearch([]int{1, 3, 5, 7}, 5)
__p(fmt.Sprint(i))
__p(fmt.Sprint(ok)) 
__check("2\ntrue")
}
