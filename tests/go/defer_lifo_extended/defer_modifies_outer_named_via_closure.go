// vybe-test: go/defer_lifo_extended/defer_modifies_outer_named_via_closure
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { result := 0
defer func() { result = 5 }()
__check(fmt.Sprint(result), "0") }
