// vybe-test: go/panic_recover_rules/defer_recover_prints_after_deferred_work
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func run() { defer __check(fmt.Sprint("late"), "saved")
defer func() { if recover() != nil { __check(fmt.Sprint("saved"), "late") } }()
panic("fail") }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { run() }
