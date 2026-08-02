// vybe-test: go/builtins_expressions_extra/make_map_then_assign_runtime
// origin: languages/go/tests/go/test_builtins_expressions_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { values := make(map[string]int)
values["go"] = 7
__check(fmt.Sprint(values["go"]), "7")
}
