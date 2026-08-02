// vybe-test: go/type_switch_extended/type_switch_assert_pointer_deref_in_body
// origin: languages/go/tests/go/test_type_switch_extended.rs

package main
import "fmt"
func work(v interface{}) { switch p := v.(type) { case *int: fmt.Println(*p)
default: fmt.Println(0) } }
func main() { n := 9
work(&n) }
