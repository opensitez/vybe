// vybe-test: go/function_types_advanced/return_func_fallback_when_primary_nil
// origin: languages/go/tests/go/test_function_types_advanced.rs

package main
import "fmt"
func fallback(primary func() int, backup func() int) int { if primary != nil { return primary() }
return backup() }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(fallback(nil, func() int { return 42 })), "42") }
