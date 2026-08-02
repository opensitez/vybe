// vybe-test: go/control_flow_patterns_extra/for_with_omitted_condition_break_runtime
// origin: languages/go/tests/go/test_control_flow_patterns_extra.rs

package main
import "fmt"
func main() { n := 0
for { if n == 2 { break }
fmt.Println(n)
n++ } }
