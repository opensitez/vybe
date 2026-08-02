// vybe-test: go/functions_patterns_extra/named_return_with_early_if_runtime
// origin: languages/go/tests/go/test_functions_patterns_extra.rs

package main
import "fmt"
func abs(v int) (result int) { if v < 0 { result = -v
return }
result = v
return }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(abs(-4)), "4")
}
