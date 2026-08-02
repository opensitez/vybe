// vybe-test: go/switch_fallthrough_extended/switch_on_named_type
// origin: languages/go/tests/go/test_switch_fallthrough_extended.rs

package main
import "fmt"
type code int
func main() { switch code(2) { case code(2): fmt.Println("c")
default: fmt.Println("d") } }
