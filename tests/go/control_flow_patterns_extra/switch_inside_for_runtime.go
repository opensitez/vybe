// vybe-test: go/control_flow_patterns_extra/switch_inside_for_runtime
// origin: languages/go/tests/go/test_control_flow_patterns_extra.rs

package main
import "fmt"
func main() { for i := 0; i < 3; i++ { switch i { case 0: fmt.Println("zero")
default: fmt.Println("other") } } }
