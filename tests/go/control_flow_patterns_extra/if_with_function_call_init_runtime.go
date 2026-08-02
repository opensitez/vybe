// vybe-test: go/control_flow_patterns_extra/if_with_function_call_init_runtime
// origin: languages/go/tests/go/test_control_flow_patterns_extra.rs

package main
import "fmt"
func value() int { return 7 }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { if n := value(); n > 5 { __check(fmt.Sprint(n), "7") } }
