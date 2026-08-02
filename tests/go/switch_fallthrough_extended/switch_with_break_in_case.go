// vybe-test: go/switch_fallthrough_extended/switch_with_break_in_case
// origin: languages/go/tests/go/test_switch_fallthrough_extended.rs

package main
import "fmt"
func main() { switch 1 { case 1: fmt.Println("a")
break
case 2: fmt.Println("b") }
fmt.Println("done") }
