// vybe-test: go/interface_assertion_extended/assert_interface_implementer_method_call
// origin: languages/go/tests/go/test_interface_assertion_extended.rs

package main
import "fmt"
type greeter interface { greet() string }
type hi struct{}
func (hi) greet() string { return "yo" }
func main() { var v interface{} = hi{}
if g, ok := v.(greeter); ok { fmt.Println(g.greet()) } else { fmt.Println("no") } }
