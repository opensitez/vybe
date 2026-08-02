// vybe-test: go/type_switch_extended/type_switch_func_branch
// origin: languages/go/tests/go/test_type_switch_extended.rs

package main
import "fmt"
func tag(v interface{}) { switch v.(type) { case func(): fmt.Println("func") default: fmt.Println("other") } }
func main() { tag(func() {}) }
