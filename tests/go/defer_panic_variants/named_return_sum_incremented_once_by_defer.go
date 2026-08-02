// vybe-test: go/defer_panic_variants/named_return_sum_incremented_once_by_defer
// origin: languages/go/tests/go/test_defer_panic_variants.rs

package main
import "fmt"
func total() (sum int) { defer func() { sum = sum + 10 }()
return 7 }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(total()), "17")
}
