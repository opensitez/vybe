// vybe-test: go/variadic_spread/variadic_interface_mixed_len
// origin: languages/go/tests/go/test_variadic_spread.rs

package main
import "fmt"
func pack(values ...interface{}) int { return len(values) }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(pack(1, "two", true)), "3")
}
