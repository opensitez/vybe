// vybe-test: go/builtins_expressions_extra/delete_then_lookup_ok_runtime
// origin: languages/go/tests/go/test_builtins_expressions_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { values := map[string]int{"a": 1}
delete(values, "a")
_, ok := values["a"]
__check(fmt.Sprint(ok), "false")
}
