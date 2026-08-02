// vybe-test: go/type_aliases/defined_type_value_receiver_returns_new_value
// origin: languages/go/tests/go/test_type_aliases.rs

package main
import "fmt"
type Offset int
func (o Offset) next() Offset { return o + 1 }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(Offset(2).next()), "3") }
