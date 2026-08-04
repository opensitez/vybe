// vybe-test: go/map_iteration_delete/map_delete_values_above_threshold_during_range
// origin: languages/go/tests/go/test_map_iteration_delete.rs

package main
import "fmt"
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

func main() { values := map[string]int{"low": 1, "mid": 5, "high": 9}
for key, value := range values { if value > 5 { delete(values, key) } }
__p(fmt.Sprint(len(values)))
total := 0
for _, value := range values { total += value }
__p(fmt.Sprint(total)) 
__check("2\n6")
}
