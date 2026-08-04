// vybe-test: go/composite_literals_extra/nested_composite_literal_three_levels_runtime
// origin: languages/go/tests/go/test_composite_literals_extra.rs

package main
import "fmt"
type cell struct { value int }
type row struct { cells []cell }
type table struct { rows []row }
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

func main() { t := table{rows: []row{{cells: []cell{{value: 8}}}}}
__p(fmt.Sprint(t.rows[0].cells[0].value))
__check("8")
}
