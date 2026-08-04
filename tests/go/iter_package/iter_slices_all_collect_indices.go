// vybe-test: go/iter_package/iter_slices_all_collect_indices
// origin: languages/go/tests/go/test_iter_package.rs

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

func main() { s := []string{"a", "b", "c"}
n := 0
for i := range slices.All(s) { if s[i] != "" { n++ } }
__p(fmt.Sprint(n)) 
__check("3")
}
