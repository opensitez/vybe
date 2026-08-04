// vybe-test: go/composite_literal_keys/nested_struct_map_array_all_keyed
// origin: languages/go/tests/go/test_composite_literal_keys.rs

package main
import "fmt"
type entry struct { scores []int }
type table struct { rows map[string]entry }
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

func main() { t := table{rows: map[string]entry{"a": {scores: []int{0: 100, 2: 300}}}}
__p(fmt.Sprint(t.rows["a"].scores[0]))
__p(fmt.Sprint(t.rows["a"].scores[2]))
__check("100\n300")
}
