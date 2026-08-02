// vybe-test: go/defer_panic_variants/recover_captures_float64_panic
// origin: languages/go/tests/go/test_defer_panic_variants.rs

package main
import "fmt"
func run() { defer func() { __check(fmt.Sprint(recover()), "2.5") }()
panic(2.5) }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { run() }
