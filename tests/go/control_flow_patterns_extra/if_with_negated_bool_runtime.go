// vybe-test: go/control_flow_patterns_extra/if_with_negated_bool_runtime
// origin: languages/go/tests/go/test_control_flow_patterns_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { ok := false
if !ok { __check(fmt.Sprint("no"), "no") } }
