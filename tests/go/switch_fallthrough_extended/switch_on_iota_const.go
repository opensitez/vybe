// vybe-test: go/switch_fallthrough_extended/switch_on_iota_const
// origin: languages/go/tests/go/test_switch_fallthrough_extended.rs

package main
import "fmt"
const ( a = iota; b; c )
func main() { switch b { case 1: fmt.Println("b")
default: fmt.Println("x") } }
