// vybe-test: go/function_types_advanced/returned_func_captures_outer_by_reference
// origin: languages/go/tests/go/test_function_types_advanced.rs

package main
import "fmt"
func counter() func() int { total := 0
return func() int { total++
return total } }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { next := counter()
__check(fmt.Sprint(next()), "1")
__check(fmt.Sprint(next()), "2") }
