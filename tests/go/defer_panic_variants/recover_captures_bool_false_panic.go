// vybe-test: go/defer_panic_variants/recover_captures_bool_false_panic
// origin: languages/go/tests/go/test_defer_panic_variants.rs

package main
import "fmt"
func run() { defer func() { value := recover()
__check(fmt.Sprint(value == false), "true") }()
panic(false) }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { run() }
