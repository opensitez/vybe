// vybe-test: go/for_range_extended/range_map_int_keys_only_max
// origin: languages/go/tests/go/test_for_range_extended.rs

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

func main() { max := 0
for k := range map[int]string{3: "c", 7: "g", 1: "a"} { if k > max { max = k } }
__p(fmt.Sprint(max)) 
__check("7")
}
