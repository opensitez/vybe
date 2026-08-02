// vybe-test: go/builtins_expressions_extra/append_result_reassigned_runtime
// origin: languages/go/tests/go/test_builtins_expressions_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { values := []int{1, 2}
values = append(values, 5)
__check(fmt.Sprint(values[2]), "5")
}
