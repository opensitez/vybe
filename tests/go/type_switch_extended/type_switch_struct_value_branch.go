// vybe-test: go/type_switch_extended/type_switch_struct_value_branch
// origin: languages/go/tests/go/test_type_switch_extended.rs

package main
import "fmt"
type box struct { n int }
func tag(v interface{}) { switch v.(type) { case box: fmt.Println("value") case *box: fmt.Println("ptr") default: fmt.Println("other") } }
func main() { tag(box{n: 2}) }
