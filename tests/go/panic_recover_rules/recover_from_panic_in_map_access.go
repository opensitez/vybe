// vybe-test: go/panic_recover_rules/recover_from_panic_in_map_access
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func run() { defer func() { __check(fmt.Sprint(recover()), "map") }()
m := map[string]int{}
_ = m["missing"]
panic("map") }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { run() }
