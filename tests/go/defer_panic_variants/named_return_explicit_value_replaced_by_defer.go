// vybe-test: go/defer_panic_variants/named_return_explicit_value_replaced_by_defer
// origin: languages/go/tests/go/test_defer_panic_variants.rs

package main
import "fmt"
func build() (n int) { defer func() { n = 99 }()
return 5 }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(build()), "99")
}
