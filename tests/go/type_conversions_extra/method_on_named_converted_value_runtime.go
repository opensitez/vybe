// vybe-test: go/type_conversions_extra/method_on_named_converted_value_runtime
// origin: languages/go/tests/go/test_type_conversions_extra.rs

package main
import "fmt"
type level int
func (l level) next() int { return int(l) + 1 }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(level(17).next()), "18")
}
