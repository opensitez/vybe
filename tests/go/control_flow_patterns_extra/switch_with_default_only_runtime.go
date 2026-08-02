// vybe-test: go/control_flow_patterns_extra/switch_with_default_only_runtime
// origin: languages/go/tests/go/test_control_flow_patterns_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { switch 99 { default: __check(fmt.Sprint("default"), "default") } }
