// vybe-test: go/defer_panic_variants/recover_from_plain_function_is_nil
// origin: languages/go/tests/go/test_defer_panic_variants.rs

package main
import "fmt"
func probe() { __check(fmt.Sprint(recover() == nil), "true") }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { probe() }
