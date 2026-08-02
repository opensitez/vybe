// vybe-test: go/control_flow_patterns_extra/for_with_post_assignment_runtime
// origin: languages/go/tests/go/test_control_flow_patterns_extra.rs

package main
import "fmt"
func main() { for i := 1; i < 8; i = i * 2 { fmt.Println(i) } }
