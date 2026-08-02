// vybe-test: go/functions_patterns_extra/function_literal_in_struct_field_runtime
// origin: languages/go/tests/go/test_functions_patterns_extra.rs

package main
import "fmt"
type holder struct { fn func(int) int }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { h := holder{fn: func(v int) int { return v * 3 }}
__check(fmt.Sprint(h.fn(4)), "12")
}
