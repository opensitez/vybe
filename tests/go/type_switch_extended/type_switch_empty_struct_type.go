// vybe-test: go/type_switch_extended/type_switch_empty_struct_type
// origin: languages/go/tests/go/test_type_switch_extended.rs

package main
import "fmt"
type empty struct{}
func tag(v interface{}) { switch v.(type) { case empty: fmt.Println("empty")
default: fmt.Println("other") } }
func main() { tag(empty{}) }
