// vybe-test: go/switch_fallthrough_extended/switch_no_fallthrough_stops
// origin: languages/go/tests/go/test_switch_fallthrough_extended.rs

package main
import "fmt"
func main() { switch 1 { case 1: fmt.Println(1)
case 2: fmt.Println(2) } }
