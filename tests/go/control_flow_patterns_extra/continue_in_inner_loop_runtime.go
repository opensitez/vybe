// vybe-test: go/control_flow_patterns_extra/continue_in_inner_loop_runtime
// origin: languages/go/tests/go/test_control_flow_patterns_extra.rs

package main
import "fmt"
func main() { for i := 0; i < 2; i++ { for j := 0; j < 3; j++ { if j == 1 { continue }
fmt.Println(i + j) } } }
