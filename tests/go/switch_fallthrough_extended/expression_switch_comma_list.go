// vybe-test: go/switch_fallthrough_extended/expression_switch_comma_list
// origin: languages/go/tests/go/test_switch_fallthrough_extended.rs

package main
import "fmt"
func main() { switch 2 { case 1, 2, 3: fmt.Println("hit")
default: fmt.Println("miss") } }
