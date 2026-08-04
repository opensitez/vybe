// vybe-test: go/sort_stable_search/sort_slice_float64_abs
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

func main() { f := []float64{-3.0, 1.0, -2.0, 4.0}
sort.Slice(f, func(i, j int) bool { ai, aj := f[i], f[j]; if ai < 0 { ai = -ai }; if aj < 0 { aj = -aj }; return ai < aj })
__p(fmt.Sprint(f[0]))
__p(fmt.Sprint(f[3])) 
__check("1\n-3")
}
