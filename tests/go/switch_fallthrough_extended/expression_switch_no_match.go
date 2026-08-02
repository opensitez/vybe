// vybe-test: go/switch_fallthrough_extended/expression_switch_no_match
// origin: languages/go/tests/go/test_switch_fallthrough_extended.rs

package main
import "fmt"
func main() { switch 5 { case 1, 2: fmt.Println("hit")
default: fmt.Println("miss") } }
