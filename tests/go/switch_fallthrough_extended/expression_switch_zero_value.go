// vybe-test: go/switch_fallthrough_extended/expression_switch_zero_value
// origin: languages/go/tests/go/test_switch_fallthrough_extended.rs

package main
import "fmt"
func main() { switch 0 { case 0: fmt.Println("zero")
default: fmt.Println("other") } }
