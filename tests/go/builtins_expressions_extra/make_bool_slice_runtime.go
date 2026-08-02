// vybe-test: go/builtins_expressions_extra/make_bool_slice_runtime
// origin: languages/go/tests/go/test_builtins_expressions_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { flags := make([]bool, 2)
__check(fmt.Sprint(len(flags)), "2")
__check(fmt.Sprint(flags[1]), "false")
}
