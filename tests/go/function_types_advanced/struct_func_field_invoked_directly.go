// vybe-test: go/function_types_advanced/struct_func_field_invoked_directly
// origin: languages/go/tests/go/test_function_types_advanced.rs

package main
import "fmt"
type worker struct { run func(int) int }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { value := worker{run: func(v int) int { return v + 5 }}
__check(fmt.Sprint(value.run(7)), "12") }
