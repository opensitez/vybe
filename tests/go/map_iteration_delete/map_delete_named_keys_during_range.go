// vybe-test: go/map_iteration_delete/map_delete_named_keys_during_range
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

func main() { values := map[string]int{"pre_a": 1, "pre_b": 2, "ok": 3}
for key := range values { if key == "pre_a" || key == "pre_b" { delete(values, key) } }
__p(fmt.Sprint(len(values)))
__p(fmt.Sprint(values["ok"])) 
__check("1\n3")
}
