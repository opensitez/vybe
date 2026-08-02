// vybe-test: go/type_switch_extended/type_switch_embedded_anonymous_struct
// origin: languages/go/tests/go/test_type_switch_extended.rs

package main
import "fmt"
type base struct { id int }
type child struct { base
name string }
func tag(v interface{}) { switch t := v.(type) { case child: fmt.Println(t.id)
fmt.Println(t.name)
default: fmt.Println("x") } }
func main() { tag(child{base: base{id: 3}, name: "c"}) }
