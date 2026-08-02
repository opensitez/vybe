// vybe-test: go/control_flow_patterns_extra/for_with_condition_only_runtime
// origin: languages/go/tests/go/test_control_flow_patterns_extra.rs

package main
import "fmt"
func main() { n := 0
for n < 3 { fmt.Println(n)
n++ } }
