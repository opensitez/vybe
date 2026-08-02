// vybe-test: go/control_flow_patterns_extra/switch_with_expressionless_true_runtime
// origin: languages/go/tests/go/test_control_flow_patterns_extra.rs

package main
import "fmt"
func main() { n := 12
switch { case n < 10: fmt.Println("small")
case n < 20: fmt.Println("medium") } }
