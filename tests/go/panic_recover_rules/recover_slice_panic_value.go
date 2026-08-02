// vybe-test: go/panic_recover_rules/recover_slice_panic_value
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func run() { defer func() { s := recover().([]int)
__check(fmt.Sprint(len(s)), "3") }()
panic([]int{1, 2, 3}) }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { run() }
