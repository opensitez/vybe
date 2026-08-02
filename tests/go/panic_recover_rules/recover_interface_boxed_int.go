// vybe-test: go/panic_recover_rules/recover_interface_boxed_int
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func run() { defer func() { v := recover().(interface{})
__check(fmt.Sprint(v.(int)), "5") }()
panic(interface{}(5)) }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { run() }
