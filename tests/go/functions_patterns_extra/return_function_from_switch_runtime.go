// vybe-test: go/functions_patterns_extra/return_function_from_switch_runtime
// origin: languages/go/tests/go/test_functions_patterns_extra.rs

package main
import "fmt"
func choose(flag bool) func(int) int { switch flag { case true: return func(v int) int { return v + 1 }
default: return func(v int) int { return v - 1 } } }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(choose(true)(8)), "9")
}
