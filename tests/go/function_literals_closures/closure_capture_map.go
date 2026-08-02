// vybe-test: go/function_literals_closures/closure_capture_map
// origin: languages/go/tests/go/test_function_literals_closures.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { m := map[string]int{"a": 1}
lookup := func(k string) int { return m[k] }
__check(fmt.Sprint(lookup("a")), "1") }
