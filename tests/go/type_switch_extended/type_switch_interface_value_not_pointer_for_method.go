// vybe-test: go/type_switch_extended/type_switch_interface_value_not_pointer_for_method
// origin: languages/go/tests/go/test_type_switch_extended.rs

package main
import "fmt"
type valuer interface { Value() int }
type thing struct { n int }
func (t thing) Value() int { return t.n }
func tag(v interface{}) { switch v.(type) { case valuer: fmt.Println(v.(valuer).Value())
default: fmt.Println(0) } }
func main() { tag(thing{n: 6}) }
