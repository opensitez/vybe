// vybe-test: go/sort_stable_search/sort_slice_by_string_length
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

func main() { s := []string{"go", "vybe", "a", "lang"}
sort.Slice(s, func(i, j int) bool { return len(s[i]) < len(s[j]) })
__p(fmt.Sprint(s[0]))
__p(fmt.Sprint(s[3])) 
__check("a\nvybe")
}
