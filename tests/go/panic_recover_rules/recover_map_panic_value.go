// vybe-test: go/panic_recover_rules/recover_map_panic_value
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func run() { defer func() { m := recover().(map[string]int)
__check(fmt.Sprint(m["k"]), "4") }()
panic(map[string]int{"k": 4}) }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { run() }
