// vybe-test: go/control_flow_patterns_extra/if_inside_switch_runtime
// origin: languages/go/tests/go/test_control_flow_patterns_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { x := 5
switch { case x > 0: if x > 3 { __check(fmt.Sprint("big"), "big") } } }
