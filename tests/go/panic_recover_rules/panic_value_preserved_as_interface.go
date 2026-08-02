// vybe-test: go/panic_recover_rules/panic_value_preserved_as_interface
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func run() { defer func() { v := recover()
__check(fmt.Sprint(v.(int) + 1), "11") }()
panic(10) }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { run() }
