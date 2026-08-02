// vybe-test: go/type_aliases/defined_int_method_on_value_receiver
// origin: languages/go/tests/go/test_type_aliases.rs

package main
import "fmt"
type Meters int
func (m Meters) Double() int { return int(m) * 2 }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(Meters(5).Double()), "10") }
