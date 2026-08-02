// vybe-test: go/type_switch_extended/type_switch_underlying_int_from_named
// origin: languages/go/tests/go/test_type_switch_extended.rs

package main
import "fmt"
type counter int
func tag(v interface{}) { switch v.(type) { case int: fmt.Println("int")
case counter: fmt.Println("counter")
default: fmt.Println("other") } }
func main() { tag(3) }
