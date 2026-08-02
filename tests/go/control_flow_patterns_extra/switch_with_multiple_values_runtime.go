// vybe-test: go/control_flow_patterns_extra/switch_with_multiple_values_runtime
// origin: languages/go/tests/go/test_control_flow_patterns_extra.rs

package main
import "fmt"
func main() { x := 3
switch x { case 1, 2: fmt.Println("low")
case 3, 4: fmt.Println("mid") } }
