// vybe-test: go/function_types_advanced/distinct_named_types_same_signature_explicit_cast
// origin: languages/go/tests/go/test_function_types_advanced.rs

package main
import "fmt"
type Adder func(int, int) int
type Combiner func(int, int) int
func use(c Combiner) int { return c(2, 5) }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var add Adder = func(a int, b int) int { return a + b }
__check(fmt.Sprint(use(Combiner(add))), "7") }
