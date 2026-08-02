// vybe-test: go/panic_recover_rules/recover_from_panic_after_fmt_sprint
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func run() { defer func() { __check(fmt.Sprint(recover()), "fmt") }()
_ = fmt.Sprint(1)
panic("fmt") }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { run() }
