// vybe-test: go/function_literals_closures/closure_stored_in_map_value
// origin: languages/go/tests/go/test_function_literals_closures.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { ops := map[string]func(int, int) int{ "add": func(a, b int) int { return a + b } }
__check(fmt.Sprint(ops["add"](3, 4)), "7") }
