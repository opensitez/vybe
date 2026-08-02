// vybe-test: go/panic_recover_rules/recover_does_not_stop_sibling_defers
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func run() { defer __check(fmt.Sprint("sibling"), "sibling")
defer func() { recover() }()
panic("p") }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { run() }
