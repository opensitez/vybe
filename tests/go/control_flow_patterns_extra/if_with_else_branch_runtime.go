// vybe-test: go/control_flow_patterns_extra/if_with_else_branch_runtime
// origin: languages/go/tests/go/test_control_flow_patterns_extra.rs

package main
import "fmt"
func main() { if value := 1; value > 3 { fmt.Println("high") } else { fmt.Println("low") } }
