// vybe-test: go/defer_panic_variants/recover_nil_panic_yields_nil
// origin: languages/go/tests/go/test_defer_panic_variants.rs

package main
import "fmt"
func run() { defer func() { __check(fmt.Sprint(recover() == nil), "true") }()
panic(nil) }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { run() }
