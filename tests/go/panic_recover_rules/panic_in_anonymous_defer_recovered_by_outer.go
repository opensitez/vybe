// vybe-test: go/panic_recover_rules/panic_in_anonymous_defer_recovered_by_outer
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func run() { defer func() { if recover() != nil { __check(fmt.Sprint("outer"), "outer") } }()
func() { panic("inner") }() }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { run() }
