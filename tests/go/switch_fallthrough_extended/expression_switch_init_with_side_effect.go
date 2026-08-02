// vybe-test: go/switch_fallthrough_extended/expression_switch_init_with_side_effect
// origin: languages/go/tests/go/test_switch_fallthrough_extended.rs

package main
import "fmt"
func main() { count := 0
switch count++; count { case 1: fmt.Println(count)
default: fmt.Println(0) } }
