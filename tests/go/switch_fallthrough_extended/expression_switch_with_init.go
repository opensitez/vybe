// vybe-test: go/switch_fallthrough_extended/expression_switch_with_init
// origin: languages/go/tests/go/test_switch_fallthrough_extended.rs

package main
import "fmt"
func main() { switch n := 4; n { case 4: fmt.Println("four")
default: fmt.Println("other") } }
