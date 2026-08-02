// vybe-test: go/defer_panic_recover_extra/recover_type_preserved_string_runtime
// origin: languages/go/tests/go/test_defer_panic_recover_extra.rs

package main
import "fmt"
func run() { defer func() { value := recover()
__check(fmt.Sprint(value == "err"), "true") }()
panic("err") }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { run() }
