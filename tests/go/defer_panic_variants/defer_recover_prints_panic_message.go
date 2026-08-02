// vybe-test: go/defer_panic_variants/defer_recover_prints_panic_message
// origin: languages/go/tests/go/test_defer_panic_variants.rs

package main
import "fmt"
func run() { defer func() { if r := recover(); r != nil { __check(fmt.Sprint(r), "halt") } }()
panic("halt") }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { run() }
