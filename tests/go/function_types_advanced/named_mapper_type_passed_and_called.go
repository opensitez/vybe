// vybe-test: go/function_types_advanced/named_mapper_type_passed_and_called
// origin: languages/go/tests/go/test_function_types_advanced.rs

package main
import "fmt"
type Mapper func(int) int
func apply(v int, m Mapper) int { return m(v) }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(apply(4, Mapper(func(v int) int { return v * 3 }))), "12") }
