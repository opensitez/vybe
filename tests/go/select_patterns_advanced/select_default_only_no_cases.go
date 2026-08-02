// vybe-test: go/select_patterns_advanced/select_default_only_no_cases
// origin: languages/go/tests/go/test_select_patterns_advanced.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { select { default: __check(fmt.Sprint("only"), "only") } }
