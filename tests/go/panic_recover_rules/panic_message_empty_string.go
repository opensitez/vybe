// vybe-test: go/panic_recover_rules/panic_message_empty_string
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func run() { defer func() { __check(fmt.Sprint(recover() == ""), "true") }()
panic("") }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { run() }
