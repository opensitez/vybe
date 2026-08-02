// vybe-test: go/control_flow_patterns_extra/for_with_multiple_init_vars_runtime
// origin: languages/go/tests/go/test_control_flow_patterns_extra.rs

package main
import "fmt"
func main() { for i, j := 0, 3; i < 3; i, j = i+1, j-1 { fmt.Println(i + j) } }
