// vybe-test: go/type_switch_extended/type_switch_assert_struct_field_in_body
// origin: languages/go/tests/go/test_type_switch_extended.rs

package main
import "fmt"
type pair struct { a int
b int }
func work(v interface{}) { switch x := v.(type) { case pair: fmt.Println(x.a + x.b)
case *pair: fmt.Println(x.b)
default: fmt.Println(-1) } }
func main() { work(pair{a: 2, b: 5}) }
