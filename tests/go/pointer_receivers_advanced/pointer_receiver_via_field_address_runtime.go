// vybe-test: go/pointer_receivers_advanced/pointer_receiver_via_field_address_runtime
// origin: languages/go/tests/go/test_pointer_receivers_advanced.rs

package main
import "fmt"
type holder struct { gauge int }
func (h *holder) raise() { h.gauge++ }
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

func main() { value := holder{gauge: 2}
alias := &value
alias.raise()
__p(fmt.Sprint(value.gauge))
__check("3")
}
