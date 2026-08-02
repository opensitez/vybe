// vybe-test: go/method_sets_pointer_value/method_on_defined_int_type_value_receiver_runtime
// origin: languages/go/tests/go/test_method_sets_pointer_value.rs

package main
import "fmt"
type counter int
func (c counter) next() int { return int(c) + 1 }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var c counter = 4
__check(fmt.Sprint(c.next()), "5") }
