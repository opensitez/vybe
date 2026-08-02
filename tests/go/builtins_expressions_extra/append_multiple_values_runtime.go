// vybe-test: go/builtins_expressions_extra/append_multiple_values_runtime
// origin: languages/go/tests/go/test_builtins_expressions_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { values := []int{1}
values = append(values, 2, 3, 4)
__check(fmt.Sprint(len(values)), "4")
__check(fmt.Sprint(values[3]), "4")
}
