// vybe-test: go/defer_panic_variants/recover_uint_panic_value
// origin: languages/go/tests/go/test_defer_panic_variants.rs

package main
import "fmt"
func run() { defer func() { __check(fmt.Sprint(recover()), "8") }()
panic(uint(8)) }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { run() }
