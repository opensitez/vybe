// vybe-test: go/control_flow_patterns_extra/break_from_switch_in_loop_runtime
// origin: languages/go/tests/go/test_control_flow_patterns_extra.rs

package main
import "fmt"
func main() { for i := 0; i < 2; i++ { switch i { case 0: fmt.Println("zero")
break
default: fmt.Println("one") } } }
