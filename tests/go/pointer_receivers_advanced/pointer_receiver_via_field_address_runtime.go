// vybe-test: go/pointer_receivers_advanced/pointer_receiver_via_field_address_runtime
// origin: languages/go/tests/go/test_pointer_receivers_advanced.rs

package main
import "fmt"
type holder struct { gauge int }
func (h *holder) raise() { h.gauge++ }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { value := holder{gauge: 2}
alias := &value
alias.raise()
__check(fmt.Sprint(value.gauge), "3")
}
