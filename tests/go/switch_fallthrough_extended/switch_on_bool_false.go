// vybe-test: go/switch_fallthrough_extended/switch_on_bool_false
// origin: languages/go/tests/go/test_switch_fallthrough_extended.rs

package main
import "fmt"
func main() { switch false { case true: fmt.Println("yes")
case false: fmt.Println("no") } }
