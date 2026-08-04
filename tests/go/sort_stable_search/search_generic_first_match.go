// vybe-test: go/sort_stable_search/search_generic_first_match
// origin: languages/go/tests/go/test_sort_stable_search.rs

package main
import "fmt"
import "sort"
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

func main() { a := []int{1, 3, 5, 7}
i := sort.Search(len(a), func(k int) bool { return a[k] >= 1 })
__p(fmt.Sprint(i)) 
__check("0")
}
