// vybe-test: go/type_conversions_extra/struct_field_conversion_runtime
// origin: languages/go/tests/go/test_type_conversions_extra.rs

package main
import "fmt"
type holder struct { count int }
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

func main() { value := holder{count: 12}
__p(fmt.Sprint(float64(value.count)))
__check("12")
}
