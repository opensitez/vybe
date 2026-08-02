// vybe-test: go/panic_recover_rules/panic_string_value_equality
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func run() { defer func() { v := recover()
__check(fmt.Sprint(v == "msg"), "true") }()
panic("msg") }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { run() }
