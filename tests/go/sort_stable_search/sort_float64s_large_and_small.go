// vybe-test: go/sort_stable_search/sort_float64s_large_and_small
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

func main() { f := []float64{1e10, 1e-10, 1.0}
sort.Float64s(f)
__p(fmt.Sprint(f[0] < f[1]))
__p(fmt.Sprint(f[2])) 
__check("true\n10000000000")
}
