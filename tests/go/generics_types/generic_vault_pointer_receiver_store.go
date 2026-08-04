// vybe-test: go/generics_types/generic_vault_pointer_receiver_store
// origin: languages/go/tests/go/test_generics_types.rs

package main
import "fmt"
type Vault[T any] struct { V T }
func (v *Vault[T]) Store(x T) { v.V = x }
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

func main() { vault := Vault[int]{V: 1}
vault.Store(99)
__p(fmt.Sprint(vault.V)) 
__check("99")
}
