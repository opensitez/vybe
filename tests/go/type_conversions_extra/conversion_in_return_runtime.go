// vybe-test: go/type_conversions_extra/conversion_in_return_runtime
// origin: languages/go/tests/go/test_type_conversions_extra.rs

package main
import "fmt"
func cast(v int) float64 { return float64(v) }
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

func main() { __p(fmt.Sprint(cast(5)))
__check("5")
}
