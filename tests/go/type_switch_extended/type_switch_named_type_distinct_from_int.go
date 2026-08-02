// vybe-test: go/type_switch_extended/type_switch_named_type_distinct_from_int
// origin: languages/go/tests/go/test_type_switch_extended.rs

package main
import "fmt"
type counter int
func tag(v interface{}) { switch v.(type) { case counter: fmt.Println("counter")
case int: fmt.Println("int")
default: fmt.Println("other") } }
func main() { tag(counter(3)) }
