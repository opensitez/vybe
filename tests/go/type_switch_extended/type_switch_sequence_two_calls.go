// vybe-test: go/type_switch_extended/type_switch_sequence_two_calls
// origin: languages/go/tests/go/test_type_switch_extended.rs

package main
import "fmt"
func tag(v interface{}) { switch v.(type) { case int: fmt.Println("i")
case string: fmt.Println("s")
default: fmt.Println("d") } }
func main() { tag(1)
tag("x") }
