// vybe-test: go/control_flow_patterns_extra/range_over_slice_index_only_runtime
// origin: languages/go/tests/go/test_control_flow_patterns_extra.rs

package main
import "fmt"
func main() { values := []int{8, 9}
for i := range values { fmt.Println(i) } }
