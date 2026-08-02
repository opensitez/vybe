// vybe-test: go/function_types_advanced/struct_field_func_cast_to_named_type
// origin: languages/go/tests/go/test_function_types_advanced.rs

package main
import "fmt"
type Mapper func(int) int
type holder struct { fn func(int) int }
func apply(v int, m Mapper) int { return m(v) }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { value := holder{fn: func(v int) int { return v + 2 }}
__check(fmt.Sprint(apply(5, Mapper(value.fn))), "7") }
