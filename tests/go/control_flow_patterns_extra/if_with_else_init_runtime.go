// vybe-test: go/control_flow_patterns_extra/if_with_else_init_runtime
// origin: languages/go/tests/go/test_control_flow_patterns_extra.rs

package main
import "fmt"
func main() { if n := 4; n%2 == 0 { fmt.Println("even") } else { fmt.Println("odd") } }
