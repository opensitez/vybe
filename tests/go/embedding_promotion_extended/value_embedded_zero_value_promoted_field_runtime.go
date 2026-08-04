// vybe-test: go/embedding_promotion_extended/value_embedded_zero_value_promoted_field_runtime
// origin: languages/go/tests/go/test_embedding_promotion_extended.rs

package main
import "fmt"
type inner struct { n int }
type outer struct { inner }
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

func main() { var o outer
__p(fmt.Sprint(o.n)) 
__check("0")
}
