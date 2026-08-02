// vybe-test: go/control_flow_patterns_extra/switch_with_tagless_boolean_cases_runtime
// origin: languages/go/tests/go/test_control_flow_patterns_extra.rs

package main
import "fmt"
func main() { n := 8
switch { case n%3 == 0: fmt.Println("three")
case n%4 == 0: fmt.Println("four") } }
