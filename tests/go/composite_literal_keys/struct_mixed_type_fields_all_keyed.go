// vybe-test: go/composite_literal_keys/struct_mixed_type_fields_all_keyed
// origin: languages/go/tests/go/test_composite_literal_keys.rs

package main
import "fmt"
type item struct { label string
count int
active bool }
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

func main() { it := item{active: true, label: "vybe", count: 7}
__p(fmt.Sprint(it.label))
__p(fmt.Sprint(it.count))
__p(fmt.Sprint(it.active))
__check("vybe\n7\ntrue")
}
