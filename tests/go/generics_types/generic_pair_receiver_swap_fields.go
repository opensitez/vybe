// vybe-test: go/generics_types/generic_pair_receiver_swap_fields
// origin: languages/go/tests/go/test_generics_types.rs

package main
import "fmt"
type Pair[T any] struct { First, Second T }
func (p *Pair[T]) Swap() { p.First, p.Second = p.Second, p.First }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { p := Pair[int]{First: 1, Second: 2}
p.Swap()
__check(fmt.Sprint(p.First), "2")
__check(fmt.Sprint(p.Second), "1") }
