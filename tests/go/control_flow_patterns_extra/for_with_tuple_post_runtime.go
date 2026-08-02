// vybe-test: go/control_flow_patterns_extra/for_with_tuple_post_runtime
// origin: languages/go/tests/go/test_control_flow_patterns_extra.rs

package main
import "fmt"
func main() { for i, j := 0, 2; i < 3; i, j = i+1, j+2 { fmt.Println(i + j) } }
