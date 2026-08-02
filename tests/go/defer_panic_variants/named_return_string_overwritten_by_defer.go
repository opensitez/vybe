// vybe-test: go/defer_panic_variants/named_return_string_overwritten_by_defer
// origin: languages/go/tests/go/test_defer_panic_variants.rs

package main
import "fmt"
func greet() (msg string) { defer func() { msg = "bye" }()
return "hi" }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(greet()), "bye")
}
