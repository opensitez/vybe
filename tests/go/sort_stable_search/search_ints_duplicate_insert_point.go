// vybe-test: go/sort_stable_search/search_ints_duplicate_insert_point
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

func main() { a := []int{1, 2, 2, 2, 5}
__p(fmt.Sprint(sort.SearchInts(a, 2)))
__p(fmt.Sprint(sort.SearchInts(a, 3))) 
__check("1\n4")
}
