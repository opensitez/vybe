// vybe-test: go/switch_fallthrough_extended/switch_multiple_cases_same_body
// origin: languages/go/tests/go/test_switch_fallthrough_extended.rs

package main
import "fmt"
func main() { switch 2 { case 1: fmt.Println("a")
case 2: fmt.Println("b")
case 3: fmt.Println("b") } }
