// vybe-test: go/switch_fallthrough_extended/expression_switch_init_shadows_outer
// origin: languages/go/tests/go/test_switch_fallthrough_extended.rs

package main
import "fmt"
func main() { n := 1
switch n := 2; n { case 2: fmt.Println(n)
default: fmt.Println(0) }
fmt.Println(n) }
