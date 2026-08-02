// vybe-test: go/control_flow_patterns_extra/range_over_nil_slice_runtime
// origin: languages/go/tests/go/test_control_flow_patterns_extra.rs

package main
import "fmt"
func main() { var values []int
for _, v := range values { fmt.Println(v) }
fmt.Println(len(values)) }
