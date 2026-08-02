// vybe-test: go/panic_recover_rules/defer_recover_with_closure_param
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func run() { defer func(label string) { if recover() != nil { __check(fmt.Sprint(label), "ok") } }("ok")
panic("x") }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { run() }
