// vybe-test: go/switch_fallthrough_extended/default_case_middle
// origin: languages/go/tests/go/test_switch_fallthrough_extended.rs

package main
import "fmt"
func main() { switch 3 { case 1: fmt.Println(1)
default: fmt.Println("mid")
case 3: fmt.Println(3) } }
