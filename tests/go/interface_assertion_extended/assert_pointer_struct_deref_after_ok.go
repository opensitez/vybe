// vybe-test: go/interface_assertion_extended/assert_pointer_struct_deref_after_ok
// origin: languages/go/tests/go/test_interface_assertion_extended.rs

package main
import "fmt"
type pair struct { a int }
func main() { p := &pair{a: 6}
var v interface{} = p
if q, ok := v.(*pair); ok { fmt.Println(q.a) } else { fmt.Println(0) } }
