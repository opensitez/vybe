// vybe-test: go/functions_patterns_extra/closure_mutates_captured_local_runtime
// origin: languages/go/tests/go/test_functions_patterns_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { total := 0
add := func(v int) { total += v }
add(2)
add(5)
__check(fmt.Sprint(total), "7")
}
