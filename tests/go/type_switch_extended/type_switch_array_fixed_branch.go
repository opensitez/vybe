// vybe-test: go/type_switch_extended/type_switch_array_fixed_branch
// origin: languages/go/tests/go/test_type_switch_extended.rs

package main
import "fmt"
func tag(v interface{}) { switch v.(type) { case [2]int: fmt.Println("array") default: fmt.Println("other") } }
func main() { tag([2]int{1, 2}) }
