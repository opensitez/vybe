// vybe-test: go/panic_recover_rules/recover_not_effective_in_same_func_as_panic
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func run() { defer func() {}()
panic("x")
_ = recover() }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { defer func() { recover() }()
defer func() { __check(fmt.Sprint("shield"), "caught") }()
defer func() { recover() }()
func() { defer func() { if recover() != nil { __check(fmt.Sprint("caught"), "shield") } }()
panic("x") }() }
