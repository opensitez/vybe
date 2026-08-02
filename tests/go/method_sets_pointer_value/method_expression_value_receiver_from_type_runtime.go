// vybe-test: go/method_sets_pointer_value/method_expression_value_receiver_from_type_runtime
// origin: languages/go/tests/go/test_method_sets_pointer_value.rs

package main
import "fmt"
type pair struct { a int }
func (p pair) first() int { return p.a }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { fn := pair.first
__check(fmt.Sprint(fn(pair{a: 8})), "8") }
