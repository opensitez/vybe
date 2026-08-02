// vybe-test: go/control_flow_patterns_extra/nested_switch_runtime
// origin: languages/go/tests/go/test_control_flow_patterns_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { x := 2
switch x { case 2: switch { case x%2 == 0: __check(fmt.Sprint("even-two"), "even-two") } } }
