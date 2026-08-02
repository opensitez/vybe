// vybe-test: go/defer_panic_variants/panic_in_helper_recovered_by_deferred_closure
// origin: languages/go/tests/go/test_defer_panic_variants.rs

package main
import "fmt"
func boom() { panic("fail") }
func run() { defer func() { if recover() != nil { __check(fmt.Sprint("saved"), "saved") } }()
boom() }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { run() }
