// vybe-test: go/functions_patterns_extra/function_literal_returns_struct_runtime
// origin: languages/go/tests/go/test_functions_patterns_extra.rs

package main
import "fmt"
type pair struct { a int
b int }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { build := func() pair { return pair{a: 3, b: 4} }
value := build()
__check(fmt.Sprint(value.a + value.b), "7")
}
