// vybe-test: go/defer_panic_variants/defer_recover_with_message_parameter
// origin: languages/go/tests/go/test_defer_panic_variants.rs

package main
import "fmt"
func run() { defer func(label string) { if recover() != nil { __check(fmt.Sprint(label), "handled") } }("handled")
panic("err") }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { run() }
