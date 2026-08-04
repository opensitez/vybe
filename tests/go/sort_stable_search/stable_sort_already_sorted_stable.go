// vybe-test: go/sort_stable_search/stable_sort_already_sorted_stable
// origin: languages/go/tests/go/test_sort_stable_search.rs

package main
import "fmt"
import "sort"
type kv struct { k, ord int }
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

func main() { s := []kv{{1, 0}, {2, 1}, {3, 2}}
sort.SliceStable(s, func(i, j int) bool { return s[i].k < s[j].k })
__p(fmt.Sprint(s[0].ord))
__p(fmt.Sprint(s[2].ord)) 
__check("0\n2")
}
