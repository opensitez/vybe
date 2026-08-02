// vybe-test: go/builtins_expressions_extra/len_after_delete_runtime
// origin: languages/go/tests/go/test_builtins_expressions_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { values := map[string]int{"a": 1, "b": 2}
delete(values, "b")
__check(fmt.Sprint(len(values)), "1")
}
