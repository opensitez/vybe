// vybe-test: go/functions_patterns_extra/function_literal_in_map_runtime
// origin: languages/go/tests/go/test_functions_patterns_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { ops := map[string]func(int) int{"inc": func(v int) int { return v + 1 }}
__check(fmt.Sprint(ops["inc"](9)), "10")
}
