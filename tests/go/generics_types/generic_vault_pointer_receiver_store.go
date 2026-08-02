// vybe-test: go/generics_types/generic_vault_pointer_receiver_store
// origin: languages/go/tests/go/test_generics_types.rs

package main
import "fmt"
type Vault[T any] struct { V T }
func (v *Vault[T]) Store(x T) { v.V = x }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { vault := Vault[int]{V: 1}
vault.Store(99)
__check(fmt.Sprint(vault.V), "99") }
