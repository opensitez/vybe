// vybe-test: go/functions_patterns_extra/named_return_bare_return_runtime
// origin: languages/go/tests/go/test_functions_patterns_extra.rs

package main
import "fmt"
func twice(v int) (result int) { result = v * 2
return }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(twice(5)), "10")
}
