// vybe-test: go/defer_panic_variants/named_return_bare_return_scaled_by_defer
// origin: languages/go/tests/go/test_defer_panic_variants.rs

package main
import "fmt"
func scale() (n int) { defer func() { n = n * 3 }()
n = 4
return }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(scale()), "12")
}
