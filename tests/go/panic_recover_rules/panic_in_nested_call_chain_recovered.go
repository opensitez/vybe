// vybe-test: go/panic_recover_rules/panic_in_nested_call_chain_recovered
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func leaf() { panic(7) }
func mid() { leaf() }
func run() { defer func() { __check(fmt.Sprint(recover()), "7") }()
mid() }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { run() }
