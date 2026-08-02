// vybe-test: go/switch_fallthrough_extended/default_case_last_position
// origin: languages/go/tests/go/test_switch_fallthrough_extended.rs

package main
import "fmt"
func main() { switch 99 { case 1: fmt.Println(1)
case 2: fmt.Println(2)
default: fmt.Println("def") } }
