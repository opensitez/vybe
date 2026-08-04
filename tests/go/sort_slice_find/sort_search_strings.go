// vybe-test: go/sort_slice_find/sort_search_strings
// origin: languages/go/tests/go/test_sort_slice_find.rs

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

func main() { s := []string{"a","c","f"}
i, ok := sort.Find(len(s), func(i int) int { if "c" < s[i] { return -1 }; if "c" > s[i] { return 1 }; return 0 })
__p(fmt.Sprint(i) + " " + fmt.Sprint(ok)) 
__check("1 true")
}
