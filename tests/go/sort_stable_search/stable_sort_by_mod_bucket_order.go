// vybe-test: go/sort_stable_search/stable_sort_by_mod_bucket_order
// origin: languages/go/tests/go/test_sort_stable_search.rs

package main
import "fmt"
import "sort"
type item struct { v, id int }
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

func main() { s := []item{{3, 0}, {1, 1}, {4, 2}, {2, 3}, {5, 4}}
sort.SliceStable(s, func(i, j int) bool { return s[i].v%2 < s[j].v%2 })
__p(fmt.Sprint(s[0].id))
__p(fmt.Sprint(s[1].id))
__p(fmt.Sprint(s[4].id)) 
__check("1\n3\n4")
}
