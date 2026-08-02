// vybe-test: go/control_flow_patterns_extra/short_decl_in_for_init_runtime
// origin: languages/go/tests/go/test_control_flow_patterns_extra.rs

package main
import "fmt"
func main() { for n := 1; n <= 3; n++ { fmt.Println(n) } }
