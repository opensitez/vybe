// vybe-test: go/generics_types/generic_counter_increment_method
// origin: languages/go/tests/go/test_generics_types.rs

package main
import "fmt"
type Counter[T ~int] struct { n T }
func (c *Counter[T]) Inc() { c.n++ }
func (c Counter[T]) Value() T { return c.n }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { c := Counter[int]{n: 4}
c.Inc()
__check(fmt.Sprint(c.Value()), "5") }
