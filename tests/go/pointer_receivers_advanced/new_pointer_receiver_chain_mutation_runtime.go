// vybe-test: go/pointer_receivers_advanced/new_pointer_receiver_chain_mutation_runtime
// origin: languages/go/tests/go/test_pointer_receivers_advanced.rs

package main
import "fmt"
type chain struct { total int }
func (c *chain) step(v int) *chain { c.total += v
return c }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { value := new(chain)
value.step(2).step(5)
__check(fmt.Sprint(value.total), "7")
}
