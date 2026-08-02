// vybe-test: go/pointer_receivers_advanced/pointer_receiver_struct_field_mutated_runtime
// origin: languages/go/tests/go/test_pointer_receivers_advanced.rs

package main
import "fmt"
type cell struct { n int }
func (c *cell) bump() { c.n++ }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { value := cell{n: 4}
value.bump()
__check(fmt.Sprint(value.n), "5")
}
