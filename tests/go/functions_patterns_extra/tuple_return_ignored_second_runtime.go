// vybe-test: go/functions_patterns_extra/tuple_return_ignored_second_runtime
// origin: languages/go/tests/go/test_functions_patterns_extra.rs

package main
import "fmt"
func pair() (int, string) { return 8, "unused" }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { value, _ := pair()
__check(fmt.Sprint(value), "8")
}
