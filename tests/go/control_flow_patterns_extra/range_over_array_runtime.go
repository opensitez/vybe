// vybe-test: go/control_flow_patterns_extra/range_over_array_runtime
// origin: languages/go/tests/go/test_control_flow_patterns_extra.rs

package main
import "fmt"
func main() { values := [3]int{5, 6, 7}
for _, v := range values { fmt.Println(v) } }
