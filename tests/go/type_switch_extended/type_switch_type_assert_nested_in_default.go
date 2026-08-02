// vybe-test: go/type_switch_extended/type_switch_type_assert_nested_in_default
// origin: languages/go/tests/go/test_type_switch_extended.rs

package main
import "fmt"
func tag(v interface{}) { switch v.(type) { case int: fmt.Println("int")
default: if s, ok := v.(string); ok { fmt.Println(s) } else { fmt.Println("none") } } }
func main() { tag("z") }
