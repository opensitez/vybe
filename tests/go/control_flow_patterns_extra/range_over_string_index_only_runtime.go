// vybe-test: go/control_flow_patterns_extra/range_over_string_index_only_runtime
// origin: languages/go/tests/go/test_control_flow_patterns_extra.rs

package main
import "fmt"
func main() { for i := range "go" { fmt.Println(i) } }
