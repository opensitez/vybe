// vybe-test: go/control_flow_patterns_extra/nested_if_with_else_if_and_init_runtime
// origin: languages/go/tests/go/test_control_flow_patterns_extra.rs

package main
import "fmt"
func main() { if n := 5; n < 0 { fmt.Println("neg") } else if n < 10 { if n%2 == 1 { fmt.Println("odd") } } }
