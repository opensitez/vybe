// vybe-test: go/functions_patterns_extra/named_func_type_runtime
// origin: languages/go/tests/go/test_functions_patterns_extra.rs

package main
import "fmt"
type op func(int, int) int
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var add op = func(a int, b int) int { return a + b }
__check(fmt.Sprint(add(2, 8)), "10")
}
