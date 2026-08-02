// vybe-test: go/panic_recover_rules/recover_pointer_panic_value
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func run() { defer func() { p := recover().(*int)
__check(fmt.Sprint(*p), "6") }()
n := 6
panic(&n) }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { run() }
