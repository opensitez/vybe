// vybe-test: go/interface_assertion_extended/assert_func_type_from_interface
// origin: languages/go/tests/go/test_interface_assertion_extended.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { fn := func(x int) int { return x + 1 }
var v interface{} = fn
f, ok := v.(func(int) int)
__check(fmt.Sprint(f(4)), "5")
__check(fmt.Sprint(ok), "true") }
