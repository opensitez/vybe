// vybe-test: go/defer_panic_variants/panic_in_inner_defer_caught_by_outer_recover
// origin: languages/go/tests/go/test_defer_panic_variants.rs

package main
import "fmt"
func run() { defer func() { if recover() != nil { __check(fmt.Sprint("caught"), "caught") } }()
defer func() { panic("inner") }() }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { run() }
