// vybe-test: go/type_switch_extended/type_switch_bool_branch
// origin: languages/go/tests/go/test_type_switch_extended.rs

package main
import "fmt"
func tag(v interface{}) { switch v.(type) { case bool: fmt.Println("bool") default: fmt.Println("other") } }
func main() { tag(true) }
