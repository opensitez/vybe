// vybe-test: go/sort_stable_search/sort_slice_struct_by_field
// origin: languages/go/tests/go/test_sort_stable_search.rs

package main
import "fmt"
import "sort"
type pair struct { k, v int }
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

func main() { s := []pair{{3, 30}, {1, 10}, {2, 20}}
sort.Slice(s, func(i, j int) bool { return s[i].k < s[j].k })
__p(fmt.Sprint(s[0].v))
__p(fmt.Sprint(s[2].v)) 
__check("10\n30")
}
